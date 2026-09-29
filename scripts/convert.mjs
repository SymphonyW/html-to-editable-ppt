#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import { pathToFileURL } from 'node:url';
import { chromium } from 'playwright';
import PptxGenJS from 'pptxgenjs';

const DEFAULTS = {
  width: 1280,
  height: 720,
  layout: 'wide',
  report: null,
  selector: null,
};

const LAYOUTS = {
  wide: { width: 13.333, height: 7.5 },
  standard: { width: 10, height: 7.5 },
};

function printHelp() {
  console.log(`\nHTML to Editable PPT\n\nUsage:\n  node scripts/convert.mjs <input.html> <output.pptx> [options]\n\nOptions:\n  --width <px>       Browser viewport width (default: 1280)\n  --height <px>      Browser viewport height (default: 720)\n  --layout <name>    wide | standard (default: wide)\n  --report <file>    Write a JSON conversion report\n  --selector <css>    Explicit CSS selector for slide containers\n  --help             Show this help\n`);
}

function parseArgs(argv) {
  if (argv.includes('--help') || argv.includes('-h')) {
    printHelp();
    process.exit(0);
  }

  const positional = [];
  const options = { ...DEFAULTS };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }

    const name = arg.slice(2);
    const value = argv[i + 1];
    if (['width', 'height', 'layout', 'report', 'selector'].includes(name)) {
      if (!value || value.startsWith('--')) throw new Error(`Missing value for --${name}`);
      options[name] = name === 'width' || name === 'height' ? Number(value) : value;
      i += 1;
    } else {
      throw new Error(`Unknown option: ${arg}`);
    }
  }

  if (positional.length < 2) {
    printHelp();
    throw new Error('input.html and output.pptx are required');
  }

  if (!LAYOUTS[options.layout]) throw new Error(`Unsupported layout: ${options.layout}`);
  if (!Number.isFinite(options.width) || options.width <= 0) throw new Error('--width must be a positive number');
  if (!Number.isFinite(options.height) || options.height <= 0) throw new Error('--height must be a positive number');

  return {
    input: path.resolve(positional[0]),
    output: path.resolve(positional[1]),
    ...options,
  };
}

function mimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return ({
    '.html': 'text/html; charset=utf-8',
    '.htm': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
    '.ttf': 'font/ttf',
    '.otf': 'font/otf',
  })[ext] ?? 'application/octet-stream';
}

function createStaticServer(rootDir) {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      try {
        const rawPath = decodeURIComponent((req.url ?? '/').split('?')[0]);
        const safePath = path.normalize(rawPath).replace(/^([/\\])+/, '');
        const filePath = path.resolve(rootDir, safePath);
        const root = path.resolve(rootDir);
        if (!(filePath === root || filePath.startsWith(`${root}${path.sep}`))) {
          res.writeHead(403).end('Forbidden');
          return;
        }
        if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
          res.writeHead(404).end('Not found');
          return;
        }
        res.setHeader('Content-Type', mimeType(filePath));
        res.setHeader('Access-Control-Allow-Origin', '*');
        fs.createReadStream(filePath).pipe(res);
      } catch (error) {
        res.writeHead(500).end(String(error));
      }
    });

    server.on('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      resolve({
        server,
        port: address.port,
      });
    });
  });
}

function cssColorToPpt(cssColor) {
  if (!cssColor || cssColor === 'transparent') return null;
  const match = cssColor.match(/rgba?\(([^)]+)\)/i);
  if (!match) return null;
  const parts = match[1].split(',').map((p) => p.trim());
  const r = Math.max(0, Math.min(255, Math.round(Number(parts[0]))));
  const g = Math.max(0, Math.min(255, Math.round(Number(parts[1]))));
  const b = Math.max(0, Math.min(255, Math.round(Number(parts[2]))));
  const alpha = parts.length >= 4 ? Math.max(0, Math.min(1, Number(parts[3]))) : 1;
  return {
    color: [r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase(),
    transparency: Math.round((1 - alpha) * 100),
    alpha,
  };
}

function localPathFromUrl(url, serverPort, rootDir) {
  try {
    const parsed = new URL(url);
    if (parsed.hostname === '127.0.0.1' && Number(parsed.port) === Number(serverPort)) {
      const pathname = decodeURIComponent(parsed.pathname).replace(/^\//, '');
      const resolved = path.resolve(rootDir, pathname);
      const root = path.resolve(rootDir);
      return (resolved === root || resolved.startsWith(`${root}${path.sep}`)) ? resolved : null;
    }
  } catch {
    return null;
  }
  return null;
}

function extensionMime(url) {
  try {
    return mimeType(new URL(url).pathname);
  } catch {
    return mimeType(url);
  }
}

async function urlToData(url, serverPort, rootDir) {
  if (!url) return null;
  if (url.startsWith('data:')) return url.slice(5);
  if (/^[a-z0-9.+-]+\/[a-z0-9.+-]+;base64,/i.test(url)) return url;

  const localPath = localPathFromUrl(url, serverPort, rootDir);
  if (localPath && fs.existsSync(localPath)) {
    const buf = fs.readFileSync(localPath);
    return `${mimeType(localPath)};base64,${buf.toString('base64')}`;
  }

  if (url.startsWith('file:')) {
    const filePath = new URL(url);
    const buf = fs.readFileSync(filePath);
    return `${mimeType(filePath.pathname)};base64,${buf.toString('base64')}`;
  }

  if (/^https?:/i.test(url)) {
    try {
      const response = await fetch(url);
      if (!response.ok) return null;
      const arrayBuffer = await response.arrayBuffer();
      const type = response.headers.get('content-type') || extensionMime(url);
      return `${type};base64,${Buffer.from(arrayBuffer).toString('base64')}`;
    } catch {
      return null;
    }
  }

  return null;
}

function encodeSvg(svg) {
  return `image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`;
}

function pxToInX(px, viewportWidth, slideWidthIn) {
  return (px / viewportWidth) * slideWidthIn;
}

function pxToInY(px, viewportHeight, slideHeightIn) {
  return (px / viewportHeight) * slideHeightIn;
}

function clampBox(box, viewport, slide) {
  const x = Math.max(0, pxToInX(box.x, viewport.width, slide.width));
  const y = Math.max(0, pxToInY(box.y, viewport.height, slide.height));
  const w = Math.min(slide.width - x, Math.max(0.001, pxToInX(box.w, viewport.width, slide.width)));
  const h = Math.min(slide.height - y, Math.max(0.001, pxToInY(box.h, viewport.height, slide.height)));
  return { x, y, w, h };
}

function cleanFontFace(value) {
  if (!value) return 'Arial';
  return value.split(',')[0].trim().replace(/^['"]|['"]$/g, '') || 'Arial';
}

function normalizedText(text) {
  return String(text ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

async function detectSlideCount(page, selector) {
  return page.evaluate(({ selector }) => {
    function revealSlides() {
      const roots = [...document.querySelectorAll('.reveal .slides > section')];
      if (!roots.length) return [];
      const flattened = [];
      for (const root of roots) {
        const nested = [...root.querySelectorAll(':scope > section')];
        if (nested.length) flattened.push(...nested);
        else flattened.push(root);
      }
      return flattened;
    }

    function commonSlides() {
      const selectors = ['.slide', '[data-slide]', '[data-slide-index]', '.page'];
      for (const selector of selectors) {
        const els = [...document.querySelectorAll(selector)];
        if (els.length > 1 || (els.length === 1 && selector !== '.page')) return els;
      }
      return [];
    }

    function bodyChildren() {
      const children = [...document.body.children].filter((el) => {
        if (['SCRIPT', 'STYLE', 'LINK'].includes(el.tagName)) return false;
        const r = el.getBoundingClientRect();
        return r.width >= window.innerWidth * 0.65 && r.height >= window.innerHeight * 0.55;
      });
      return children.length > 1 ? children : [];
    }

    if (selector) {
      const explicit = [...document.querySelectorAll(selector)];
      return Math.max(1, explicit.length);
    }
    const slides = revealSlides().length ? revealSlides() : (commonSlides().length ? commonSlides() : bodyChildren());
    return Math.max(1, slides.length);
  }, { selector });
}

async function extractSlide(page, slideIndex, selector) {
  return page.evaluate(async ({ slideIndex, selector }) => {
    function getSlides() {
      if (selector) {
        const explicit = [...document.querySelectorAll(selector)];
        if (explicit.length) return explicit;
      }
      const roots = [...document.querySelectorAll('.reveal .slides > section')];
      if (roots.length) {
        const flattened = [];
        for (const root of roots) {
          const nested = [...root.querySelectorAll(':scope > section')];
          if (nested.length) flattened.push(...nested);
          else flattened.push(root);
        }
        return flattened;
      }

      for (const selector of ['.slide', '[data-slide]', '[data-slide-index]', '.page']) {
        const els = [...document.querySelectorAll(selector)];
        if (els.length > 1 || (els.length === 1 && selector !== '.page')) return els;
      }

      const children = [...document.body.children].filter((el) => {
        if (['SCRIPT', 'STYLE', 'LINK'].includes(el.tagName)) return false;
        const r = el.getBoundingClientRect();
        return r.width >= window.innerWidth * 0.65 && r.height >= window.innerHeight * 0.55;
      });
      if (children.length > 1) return children;
      return [document.body];
    }

    function forceSlideVisible(target, slides) {
      const ancestorSlides = new Set();
      let p = target.parentElement;
      while (p) {
        if (slides.includes(p)) ancestorSlides.add(p);
        p = p.parentElement;
      }

      for (const slide of slides) {
        if (slide === target || ancestorSlides.has(slide)) continue;
        slide.style.setProperty('display', 'none', 'important');
      }

      for (const el of [...ancestorSlides, target]) {
        el.style.setProperty('display', 'block', 'important');
        el.style.setProperty('visibility', 'visible', 'important');
        el.style.setProperty('opacity', '1', 'important');
        el.style.setProperty('transform', 'none', 'important');
      }

      target.style.setProperty('position', 'absolute', 'important');
      target.style.setProperty('left', '0', 'important');
      target.style.setProperty('top', '0', 'important');
      target.style.setProperty('width', `${window.innerWidth}px`, 'important');
      target.style.setProperty('height', `${window.innerHeight}px`, 'important');
      target.style.setProperty('margin', '0', 'important');
      target.style.setProperty('overflow', 'hidden', 'important');
      target.style.setProperty('transform', 'none', 'important');
    }

    function visible(el, slideRect) {
      const s = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      if (s.display === 'none' || s.visibility === 'hidden') return false;
      if (Number.parseFloat(s.opacity || '1') <= 0.01) return false;
      if (r.width < 0.5 || r.height < 0.5) return false;
      const intersects = r.right > slideRect.left && r.bottom > slideRect.top && r.left < slideRect.right && r.top < slideRect.bottom;
      return intersects;
    }

    function box(el, slideRect) {
      const r = el.getBoundingClientRect();
      return {
        x: r.left - slideRect.left,
        y: r.top - slideRect.top,
        w: r.width,
        h: r.height,
      };
    }

    function zValue(style) {
      const z = Number.parseInt(style.zIndex, 10);
      return Number.isFinite(z) ? z : 0;
    }

    function hasMeaningfulBackground(style) {
      const bg = style.backgroundColor;
      const rgba = bg.match(/rgba?\(([^)]+)\)/i);
      if (rgba) {
        const parts = rgba[1].split(',').map((p) => p.trim());
        const alpha = parts.length >= 4 ? Number(parts[3]) : 1;
        if (alpha > 0.01) return true;
      }
      return style.backgroundImage && style.backgroundImage !== 'none';
    }

    // Approximate a CSS gradient with the alpha-weighted average of its color
    // stops, returned as an rgba() string the PPT color parser understands.
    // Transparent stops (browsers serialize `transparent` as rgba(0,0,0,0))
    // lower the average alpha instead of darkening the color.
    function gradientApproxColor(backgroundImage) {
      const value = backgroundImage || '';
      if (!/gradient\(/i.test(value)) return null;
      const colors = [];
      const hexRe = /#([0-9a-f]{3,8})\b/gi;
      let m;
      while ((m = hexRe.exec(value))) {
        let h = m[1];
        let alpha = 1;
        if (h.length === 3 || h.length === 4) {
          if (h.length === 4) alpha = parseInt(h[3] + h[3], 16) / 255;
          h = h.slice(0, 3).split('').map((c) => c + c).join('');
        } else {
          if (h.length === 8) alpha = parseInt(h.slice(6, 8), 16) / 255;
          h = h.slice(0, 6);
        }
        colors.push([parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16), alpha]);
      }
      const rgbRe = /rgba?\(([^)]+)\)/gi;
      while ((m = rgbRe.exec(value))) {
        const p = m[1].split(',').map((s) => Number.parseFloat(s));
        if (p.length >= 3 && p.slice(0, 3).every((v) => Number.isFinite(v))) {
          colors.push([p[0], p[1], p[2], p.length >= 4 && Number.isFinite(p[3]) ? p[3] : 1]);
        }
      }
      if (!colors.length) return null;
      const weightSum = colors.reduce((acc, c) => acc + c[3], 0);
      if (weightSum <= 0) return null;
      const avg = [0, 1, 2].map((i) => Math.round(colors.reduce((acc, c) => acc + c[i] * c[3], 0) / weightSum));
      const alpha = colors.reduce((acc, c) => acc + c[3], 0) / colors.length;
      return `rgba(${avg[0]}, ${avg[1]}, ${avg[2]}, ${Math.round(alpha * 1000) / 1000})`;
    }

    function backgroundUrl(style) {
      const value = style.backgroundImage || '';
      const match = value.match(/url\([\"']?(.*?)[\"']?\)/i);
      if (!match?.[1]) return null;
      try { return new URL(match[1], document.baseURI).href; } catch { return match[1]; }
    }

    function hasBorder(style) {
      return ['Top', 'Right', 'Bottom', 'Left'].some((side) => {
        return Number.parseFloat(style[`border${side}Width`] || '0') > 0 && style[`border${side}Style`] !== 'none';
      });
    }

    // display:block/flex/... makes even a SPAN lay out as a block; tag name
    // alone is not enough to tell text containers from layout containers.
    function isBlockLike(el) {
      return /^(block|flex|grid|table|list-item|flow-root)/.test(getComputedStyle(el).display);
    }

    function directOrSimpleText(el) {
      const allowed = new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'LI', 'BLOCKQUOTE', 'PRE', 'CODE', 'TD', 'TH', 'FIGCAPTION', 'CAPTION', 'BUTTON', 'A', 'LABEL', 'DIV', 'SPAN', 'B', 'STRONG', 'EM', 'I', 'U', 'S', 'SMALL', 'MARK', 'SUB', 'SUP', 'FONT']);
      if (!allowed.has(el.tagName)) return false;
      // A container with block-level children is a layout node, not a text
      // node: merging it would stack distinct children into one text box and
      // lose their individual positions and colors.
      const blockTags = new Set(['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TABLE', 'SECTION', 'ARTICLE', 'PRE', 'BLOCKQUOTE']);
      return ![...el.children].some((child) => blockTags.has(child.tagName) || isBlockLike(child));
    }

    // Walk a text container and split its content into runs that share the
    // same inline style, so per-span colors/weights survive the conversion.
    function collectRuns(el) {
      const runs = [];
      function styleOf(node) {
        const cs = getComputedStyle(node);
        return {
          color: effectiveTextColor(cs),
          fontFamily: cs.fontFamily,
          fontSize: Number.parseFloat(cs.fontSize || '16'),
          fontWeight: cs.fontWeight,
          fontStyle: cs.fontStyle,
          textDecoration: cs.textDecorationLine || cs.textDecoration,
        };
      }
      function sameStyle(a, b) {
        return a && b && a.color === b.color && a.fontSize === b.fontSize && a.fontWeight === b.fontWeight
          && a.fontStyle === b.fontStyle && a.textDecoration === b.textDecoration && a.fontFamily === b.fontFamily;
      }
      function push(text, style) {
        if (!text) return;
        const last = runs[runs.length - 1];
        if (last && !last.br && sameStyle(last.style, style)) last.text += text;
        else runs.push({ text, style });
      }
      let lastRect = null;
      function walk(node, style) {
        if (node.nodeType === Node.TEXT_NODE) {
          push(node.nodeValue.replace(/ /g, ' ').replace(/\s+/g, ' '), style);
          return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return;
        const cs = getComputedStyle(node);
        if (cs.display === 'none' || cs.visibility === 'hidden' || Number.parseFloat(cs.opacity || '1') < 0.05) return;
        const st = styleOf(node);
        if (node.tagName === 'BR') { runs.push({ text: '\n', style: st, br: true }); return; }
        // Margin / flex-gap separation between inline runs is visible in the
        // browser but invisible in concatenated text; approximate with a space.
        if (runs.length && !runs[runs.length - 1].br && !/\s$/.test(runs[runs.length - 1].text)) {
          const ml = Number.parseFloat(cs.marginLeft || '0') || 0;
          let gap = 0;
          const pe = node.parentElement;
          if (pe && /flex/.test(getComputedStyle(pe).display)) {
            gap = Number.parseFloat(getComputedStyle(pe).columnGap || '0') || 0;
          }
          if (ml + gap > 3) push(' ', style);
        }
        // innerText breaks lines around block-level descendants; mirror that
        // so the runs' text stays aligned with the container's innerText.
        // Flex items are blockified by CSS (inline-block computes to block)
        // yet still render on the same line — compare rects with the element
        // walked just before to tell a real line break from blockification.
        const rect = node.getBoundingClientRect();
        const sameLine = lastRect && rect.height > 0
          && Math.abs(rect.top - lastRect.top) < Math.max(2, rect.height * 0.5);
        const blockBreak = node !== el && /^(block|flex|grid|table|list-item|flow-root)/.test(cs.display) && !sameLine;
        lastRect = rect;
        if (blockBreak && runs.length && !runs[runs.length - 1].br) runs.push({ text: '\n', style: st, br: true });
        for (const child of node.childNodes) walk(child, st);
        if (blockBreak) runs.push({ text: '\n', style: st, br: true });
      }
      for (const child of el.childNodes) walk(child, styleOf(el));
      if (runs.length) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }
      return runs.filter((r) => r.br || r.text);
    }

    // Text painted with a gradient (background-clip:text) reports a
    // transparent color; approximate it with the average gradient stop.
    function effectiveTextColor(cs) {
      const transparent = !cs.color || cs.color === 'transparent' || cs.color === 'rgba(0, 0, 0, 0)';
      if (transparent && (cs.webkitBackgroundClip === 'text' || cs.backgroundClip === 'text')) {
        return gradientApproxColor(cs.backgroundImage) || cs.color;
      }
      return cs.color;
    }

    function hasTextAncestor(el, selected) {
      let parent = el.parentElement;
      while (parent) {
        if (selected.has(parent)) return true;
        parent = parent.parentElement;
      }
      return false;
    }

    const slides = getSlides();
    const target = slides[Math.min(slideIndex, slides.length - 1)] || document.body;
    // Decks that compute fit-to-slide transforms on activation (e.g. a global
    // goToSlide) must be activated through their own navigation hook first,
    // otherwise JS-built layouts are measured in their pre-fit state.
    if (typeof window.goToSlide === 'function') {
      try {
        window.goToSlide(slideIndex + 1);
        // fit-to-slide transforms are often applied from rAF/setTimeout hooks
        await new Promise((resolve) => setTimeout(resolve, 350));
      } catch { /* not a real hook */ }
    }
    forceSlideVisible(target, slides);

    const slideRect = target.getBoundingClientRect();
    const all = [target, ...target.querySelectorAll('*')];
    const domOrder = new Map(all.map((el, i) => [el, i]));
    const objects = [];
    const unsupported = [];

    const targetStyle = getComputedStyle(target);
    const bodyStyle = getComputedStyle(document.body);
    const targetGradientColor = gradientApproxColor(targetStyle.backgroundImage);
    const backgroundColor = targetStyle.backgroundColor !== 'rgba(0, 0, 0, 0)' && targetStyle.backgroundColor !== 'transparent'
      ? targetStyle.backgroundColor
      : targetGradientColor || bodyStyle.backgroundColor;

    const rootBackgroundUrl = backgroundUrl(targetStyle) || backgroundUrl(bodyStyle);
    if (rootBackgroundUrl) {
      objects.push({
        kind: 'backgroundImage',
        box: { x: 0, y: 0, w: slideRect.width, h: slideRect.height },
        z: -1000000,
        order: -1,
        src: rootBackgroundUrl,
        opacity: 1,
      });
    } else if ((targetStyle.backgroundImage && targetStyle.backgroundImage !== 'none') || (bodyStyle.backgroundImage && bodyStyle.backgroundImage !== 'none')) {
      unsupported.push({ type: 'slide-background-image', value: targetStyle.backgroundImage || bodyStyle.backgroundImage });
    }

    // Simple backgrounds and borders become native shapes.
    for (const el of all) {
      if (el === target || ['IMG', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME'].includes(el.tagName)) continue;
      if (!visible(el, slideRect)) continue;
      const style = getComputedStyle(el);
      const bg = hasMeaningfulBackground(style);
      const border = hasBorder(style);
      if (!bg && !border) continue;

      if (style.backgroundImage && style.backgroundImage !== 'none') {
        const bgUrl = backgroundUrl(style);
        if (bgUrl) {
          objects.push({
            kind: 'backgroundImage',
            box: box(el, slideRect),
            z: zValue(style),
            order: (domOrder.get(el) ?? 0) + 0.1,
            src: bgUrl,
            opacity: Number.parseFloat(style.opacity || '1'),
          });
        } else {
          unsupported.push({ type: 'background-image', tag: el.tagName, value: style.backgroundImage.slice(0, 160) });
        }
      }

      const transparentBg = style.backgroundColor === 'rgba(0, 0, 0, 0)' || style.backgroundColor === 'transparent';
      const gradientFill = transparentBg ? gradientApproxColor(style.backgroundImage) : null;

      objects.push({
        kind: 'shape',
        box: box(el, slideRect),
        z: zValue(style),
        order: domOrder.get(el) ?? 0,
        style: {
          backgroundColor: gradientFill || style.backgroundColor,
          borderColor: style.borderTopColor,
          borderWidth: Number.parseFloat(style.borderTopWidth || '0'),
          borderStyle: style.borderTopStyle,
          borderRadius: Number.parseFloat(style.borderTopLeftRadius || '0'),
          opacity: Number.parseFloat(style.opacity || '1'),
        },
      });
    }

    // Native text objects. Prefer outer semantic text containers to avoid duplication.
    const selectedText = new Set();
    for (const el of all) {
      if (el === target || !visible(el, slideRect) || !directOrSimpleText(el) || hasTextAncestor(el, selectedText)) continue;
      let text = (el.innerText || '').replace(/\u00a0/g, ' ').trim();
      if (!text) continue;
      const style = getComputedStyle(el);
      if (el.tagName === 'LI' && !['none', ''].includes(style.listStyleType)) {
        const parentTag = el.parentElement?.tagName;
        if (parentTag === 'UL') text = `• ${text}`;
        else if (parentTag === 'OL') {
          const siblings = [...el.parentElement.children].filter((c) => c.tagName === 'LI');
          text = `${siblings.indexOf(el) + 1}. ${text}`;
        }
      }

      selectedText.add(el);
      const elBox = box(el, slideRect);
      const fontSizePx = Number.parseFloat(style.fontSize || '16');
      const lineHeightPx = Number.parseFloat(style.lineHeight) || fontSizePx * 1.2;
      const anchor = el.closest('a[href]');
      objects.push({
        kind: 'text',
        box: elBox,
        z: zValue(style),
        order: domOrder.get(el) ?? 0,
        text,
        runs: collectRuns(el),
        href: anchor?.href || null,
        singleLine: elBox.h <= lineHeightPx * 1.7,
        style: {
          color: effectiveTextColor(style),
          fontFamily: style.fontFamily,
          fontSize: fontSizePx,
          fontWeight: style.fontWeight,
          fontStyle: style.fontStyle,
          textDecoration: style.textDecorationLine || style.textDecoration,
          textAlign: style.textAlign,
          lineHeight: style.lineHeight,
          lineHeightPx,
          letterSpacing: style.letterSpacing,
          opacity: Number.parseFloat(style.opacity || '1'),
          whiteSpace: style.whiteSpace,
        },
      });
    }

    // Text nodes sitting directly inside a layout node (mixed with block-level
    // children) belong to no selected text container — extract them via Range
    // so they are not dropped.
    for (const el of all) {
      if (el === target || selectedText.has(el) || hasTextAncestor(el, selectedText)) continue;
      if (!visible(el, slideRect)) continue;
      const textNodes = [...el.childNodes].filter(
        (n) => n.nodeType === Node.TEXT_NODE && n.nodeValue.replace(/ /g, ' ').trim(),
      );
      if (!textNodes.length) continue;
      const range = document.createRange();
      let rect = null;
      for (const tn of textNodes) {
        range.selectNodeContents(tn);
        const r = range.getBoundingClientRect();
        if (r.width < 0.5 || r.height < 0.5) continue;
        rect = rect
          ? {
              left: Math.min(rect.left, r.left), top: Math.min(rect.top, r.top),
              right: Math.max(rect.right, r.right), bottom: Math.max(rect.bottom, r.bottom),
            }
          : { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
      }
      if (!rect) continue;
      const style = getComputedStyle(el);
      const text = textNodes.map((n) => n.nodeValue.replace(/ /g, ' ').replace(/\s+/g, ' ')).join(' ').trim();
      if (!text) continue;
      const fontSizePx = Number.parseFloat(style.fontSize || '16');
      const lineHeightPx = Number.parseFloat(style.lineHeight) || fontSizePx * 1.2;
      const h = rect.bottom - rect.top;
      objects.push({
        kind: 'text',
        box: { x: rect.left - slideRect.left, y: rect.top - slideRect.top, w: rect.right - rect.left, h },
        z: zValue(style),
        order: (domOrder.get(el) ?? 0) + 0.05,
        text,
        runs: [{
          text,
          style: {
            color: effectiveTextColor(style), fontFamily: style.fontFamily, fontSize: fontSizePx,
            fontWeight: style.fontWeight, fontStyle: style.fontStyle,
            textDecoration: style.textDecorationLine || style.textDecoration,
          },
        }],
        href: null,
        singleLine: h <= lineHeightPx * 1.7,
        style: {
          color: effectiveTextColor(style),
          fontFamily: style.fontFamily,
          fontSize: fontSizePx,
          fontWeight: style.fontWeight,
          fontStyle: style.fontStyle,
          textDecoration: style.textDecorationLine || style.textDecoration,
          textAlign: style.textAlign,
          lineHeight: style.lineHeight,
          lineHeightPx,
          letterSpacing: style.letterSpacing,
          opacity: Number.parseFloat(style.opacity || '1'),
          whiteSpace: style.whiteSpace,
        },
      });
    }

    // Raster images remain separate PowerPoint image objects.
    for (const el of target.querySelectorAll('img')) {
      if (!visible(el, slideRect)) continue;
      const style = getComputedStyle(el);
      objects.push({
        kind: 'image',
        box: box(el, slideRect),
        z: zValue(style),
        order: domOrder.get(el) ?? 0,
        src: el.currentSrc || el.src || el.getAttribute('src'),
        opacity: Number.parseFloat(style.opacity || '1'),
      });
    }

    // Top-level SVGs are preserved as vector image objects.
    for (const el of target.querySelectorAll('svg')) {
      if (el.closest('svg') !== el || !visible(el, slideRect)) continue;
      const style = getComputedStyle(el);
      const clone = el.cloneNode(true);
      if (!clone.getAttribute('xmlns')) clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      // PowerPoint sizes SVG images by their intrinsic width/height attributes;
      // pin them to the on-screen box so CSS-sized SVGs don't render oversized.
      const svgRect = el.getBoundingClientRect();
      clone.setAttribute('width', String(Math.max(1, Math.round(svgRect.width))));
      clone.setAttribute('height', String(Math.max(1, Math.round(svgRect.height))));
      objects.push({
        kind: 'svg',
        box: box(el, slideRect),
        z: zValue(style),
        order: domOrder.get(el) ?? 0,
        svg: clone.outerHTML,
        opacity: Number.parseFloat(style.opacity || '1'),
      });
    }

    // Canvas is an intentionally local raster fallback.
    for (const el of target.querySelectorAll('canvas')) {
      if (!visible(el, slideRect)) continue;
      const style = getComputedStyle(el);
      let data = null;
      try { data = el.toDataURL('image/png'); } catch { /* tainted canvas */ }
      objects.push({
        kind: 'canvas',
        box: box(el, slideRect),
        z: zValue(style),
        order: domOrder.get(el) ?? 0,
        data: data?.startsWith('data:') ? data.slice(5) : data,
        opacity: Number.parseFloat(style.opacity || '1'),
      });
      unsupported.push({ type: 'canvas-rasterized', tag: 'CANVAS' });
    }

    for (const el of target.querySelectorAll('iframe, video')) {
      if (!visible(el, slideRect)) continue;
      unsupported.push({ type: el.tagName.toLowerCase(), tag: el.tagName, note: 'Embedded media requires manual/local-region fallback.' });
    }

    for (const el of all) {
      if (!visible(el, slideRect)) continue;
      const before = getComputedStyle(el, '::before');
      const after = getComputedStyle(el, '::after');
      for (const [pseudo, ps] of [['::before', before], ['::after', after]]) {
        const content = ps.content;
        if (content && content !== 'none' && content !== 'normal' && content !== '""') {
          unsupported.push({ type: 'pseudo-element', tag: el.tagName, pseudo, content: content.slice(0, 120) });
        }
      }
    }

    return {
      title: document.title || `Slide ${slideIndex + 1}`,
      backgroundColor,
      objects,
      unsupported,
      slideIndex,
      sourceBox: { width: slideRect.width, height: slideRect.height },
    };
  }, { slideIndex, selector });
}

function addShape(pptx, slide, obj, box) {
  const fill = cssColorToPpt(obj.style.backgroundColor);
  const line = cssColorToPpt(obj.style.borderColor);
  const radius = Number(obj.style.borderRadius || 0);
  const type = radius >= 5 ? pptx.ShapeType.roundRect : pptx.ShapeType.rect;
  const opacity = Math.max(0, Math.min(1, Number(obj.style.opacity ?? 1)));

  slide.addShape(type, {
    ...box,
    fill: fill && fill.alpha > 0.01
      ? { color: fill.color, transparency: Math.min(100, fill.transparency + Math.round((1 - opacity) * 100)) }
      : { color: 'FFFFFF', transparency: 100 },
    line: line && Number(obj.style.borderWidth || 0) > 0 && obj.style.borderStyle !== 'none'
      ? {
          color: line.color,
          transparency: line.transparency,
          width: Math.max(0.25, Number(obj.style.borderWidth || 1) * 0.75),
        }
      : { color: 'FFFFFF', transparency: 100 },
    rectRadius: radius > 0 ? Math.min(1, radius / Math.max(1, Math.min(obj.box.w, obj.box.h))) : undefined,
  });
}

function runOptions(runStyle) {
  const color = cssColorToPpt(runStyle.color);
  const weight = String(runStyle.fontWeight || '400');
  const numericWeight = Number.parseInt(weight, 10);
  return {
    color: color?.color,
    transparency: color ? color.transparency : undefined,
    bold: weight === 'bold' || (Number.isFinite(numericWeight) && numericWeight >= 600),
    italic: String(runStyle.fontStyle).includes('italic'),
    underline: String(runStyle.textDecoration).includes('underline') ? { style: 'sng' } : undefined,
    fontSize: Math.max(1, Number(runStyle.fontSize || 16) * 0.75),
    fontFace: cleanFontFace(runStyle.fontFamily),
  };
}

function addText(slide, obj, box) {
  const color = cssColorToPpt(obj.style.color) ?? { color: '000000', transparency: 0 };
  const fontSize = Math.max(1, Number(obj.style.fontSize || 16) * 0.75);
  const text = normalizedText(obj.text);
  if (!text) return;

  const options = {
    ...box,
    margin: 0,
    fontFace: cleanFontFace(obj.style.fontFamily),
    fontSize,
    color: color.color,
    transparency: color.transparency,
    align: ['left', 'center', 'right', 'justify'].includes(obj.style.textAlign) ? obj.style.textAlign : 'left',
    valign: 'middle',
    fit: 'shrink',
    autoFit: false,
    wrap: obj.singleLine === false,
    isTextBox: true,
    paraSpaceAfter: 0,
    lineSpacing: Number.isFinite(obj.style.lineHeightPx) && obj.style.lineHeightPx > 0
      ? Math.round(obj.style.lineHeightPx * 0.75 * 100) / 100
      : undefined,
    lang: 'zh-CN',
  };
  const baseRun = runOptions(obj.style);
  Object.assign(options, { bold: baseRun.bold, italic: baseRun.italic, underline: baseRun.underline });

  // Rich-text runs preserve per-span colors and weights from the source.
  const rich = [];
  if (Array.isArray(obj.runs)) {
    for (const run of obj.runs) {
      if (run.br) {
        if (rich.length) rich[rich.length - 1].options.breakLine = true;
        continue;
      }
      if (!run.text) continue;
      rich.push({ text: run.text, options: runOptions(run.style) });
    }
  }
  const richText = (obj.runs || []).map((r) => r.text).join('').trim();
  // innerText inserts line breaks around CSS boxes (e.g. inline-flex) that the
  // run walker deliberately keeps inline, so compare with whitespace stripped:
  // a whitespace-only mismatch must not fall back to the stacked innerText.
  const useRich = rich.length >= 1 && richText
    && (normalizedText(richText) === text || richText.replace(/\s+/g, '') === text.replace(/\s+/g, ''));
  if (process.env.DEBUG_RUNS && text.includes('六维度')) {
    console.log('DEBUG', JSON.stringify({ text, richText: normalizedText(richText), runs: obj.runs, useRich }, null, 1));
  }

  const hyperlink = obj.href && /^https?:|^mailto:|^tel:/i.test(obj.href) ? { url: obj.href } : null;
  if (useRich) {
    if (hyperlink) rich.forEach((r) => { r.options.hyperlink = hyperlink; });
    slide.addText(rich, options);
  } else if (hyperlink) {
    slide.addText([{ text, options: { hyperlink } }], options);
  } else {
    slide.addText(text, options);
  }
}

async function addImage(slide, obj, box, serverPort, rootDir) {
  const data = await urlToData(obj.src, serverPort, rootDir);
  if (!data) return false;
  slide.addImage({ data, ...box, transparency: Math.round((1 - Math.max(0, Math.min(1, obj.opacity ?? 1))) * 100) });
  return true;
}

function addSvg(slide, obj, box) {
  if (!obj.svg) return false;
  slide.addImage({ data: encodeSvg(obj.svg), ...box, transparency: Math.round((1 - Math.max(0, Math.min(1, obj.opacity ?? 1))) * 100) });
  return true;
}

function addCanvas(slide, obj, box) {
  if (!obj.data) return false;
  slide.addImage({ data: obj.data, ...box, transparency: Math.round((1 - Math.max(0, Math.min(1, obj.opacity ?? 1))) * 100) });
  return true;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!fs.existsSync(args.input)) throw new Error(`Input file not found: ${args.input}`);
  fs.mkdirSync(path.dirname(args.output), { recursive: true });

  const inputDir = path.dirname(args.input);
  const cwd = path.resolve(process.cwd());
  const serverRoot = args.input.startsWith(`${cwd}${path.sep}`) ? cwd : inputDir;
  const entryFile = path.relative(serverRoot, args.input).split(path.sep).map(encodeURIComponent).join('/');
  const { server, port } = await createStaticServer(serverRoot);
  const browser = await chromium.launch({ headless: true });
  const slideSize = LAYOUTS[args.layout];
  const viewport = { width: args.width, height: args.height };
  const report = {
    source: args.input,
    output: args.output,
    layout: args.layout,
    viewport,
    selector: args.selector,
    createdAt: new Date().toISOString(),
    slides: [],
    totals: { slides: 0, shapes: 0, text: 0, images: 0, svg: 0, canvas: 0, unsupported: 0, failedImages: 0 },
  };

  try {
    const url = `http://127.0.0.1:${port}/${entryFile}`;
    const countPage = await browser.newPage({ viewport });
    await countPage.goto(url, { waitUntil: 'load' });
    await countPage.evaluate(async () => {
      if (document.fonts?.ready) await document.fonts.ready;
    });
    const slideCount = await detectSlideCount(countPage, args.selector);
    await countPage.close();

    const extractedSlides = [];
    for (let i = 0; i < slideCount; i += 1) {
      const page = await browser.newPage({ viewport });
      await page.goto(url, { waitUntil: 'load' });
      await page.evaluate(async () => {
        if (document.fonts?.ready) await document.fonts.ready;
      });
      const extracted = await extractSlide(page, i, args.selector);
      extractedSlides.push(extracted);
      await page.close();
    }

    const pptx = new PptxGenJS();
    pptx.layout = args.layout === 'standard' ? 'LAYOUT_4X3' : 'LAYOUT_WIDE';
    pptx.author = 'html-to-editable-ppt Skill';
    pptx.subject = 'Editable PowerPoint converted from HTML';
    pptx.title = path.basename(args.input);
    pptx.company = 'OpenAI-compatible Skill';
    pptx.lang = 'zh-CN';
    pptx.theme = {
      headFontFace: 'Arial',
      bodyFontFace: 'Arial',
      lang: 'zh-CN',
    };

    for (const extracted of extractedSlides) {
      const slide = pptx.addSlide();
      const bg = cssColorToPpt(extracted.backgroundColor);
      if (bg && bg.alpha > 0.01) slide.background = { color: bg.color, transparency: bg.transparency };

      const kindRank = { shape: 0, backgroundImage: 0.5, image: 1, svg: 1, canvas: 1, text: 2 };
      const objects = [...extracted.objects].sort((a, b) => {
        if (a.z !== b.z) return a.z - b.z;
        if (a.order !== b.order) return a.order - b.order;
        return (kindRank[a.kind] ?? 9) - (kindRank[b.kind] ?? 9);
      });

      const stats = { index: extracted.slideIndex + 1, shapes: 0, text: 0, images: 0, svg: 0, canvas: 0, unsupported: extracted.unsupported, failedImages: [] };

      for (const obj of objects) {
        const box = clampBox(obj.box, extracted.sourceBox || viewport, slideSize);
        if (box.w <= 0.001 || box.h <= 0.001) continue;
        try {
          if (obj.kind === 'shape') {
            addShape(pptx, slide, obj, box);
            stats.shapes += 1;
          } else if (obj.kind === 'text') {
            addText(slide, obj, box);
            stats.text += 1;
          } else if (obj.kind === 'image' || obj.kind === 'backgroundImage') {
            const ok = await addImage(slide, obj, box, port, serverRoot);
            if (ok) stats.images += 1;
            else stats.failedImages.push(obj.src);
          } else if (obj.kind === 'svg') {
            if (addSvg(slide, obj, box)) stats.svg += 1;
          } else if (obj.kind === 'canvas') {
            if (addCanvas(slide, obj, box)) stats.canvas += 1;
          }
        } catch (error) {
          stats.unsupported.push({ type: 'conversion-error', kind: obj.kind, error: String(error) });
        }
      }

      report.slides.push(stats);
      report.totals.shapes += stats.shapes;
      report.totals.text += stats.text;
      report.totals.images += stats.images;
      report.totals.svg += stats.svg;
      report.totals.canvas += stats.canvas;
      report.totals.unsupported += stats.unsupported.length;
      report.totals.failedImages += stats.failedImages.length;
    }

    report.totals.slides = extractedSlides.length;
    await pptx.writeFile({ fileName: args.output });

    if (args.report) {
      const reportPath = path.resolve(args.report);
      fs.mkdirSync(path.dirname(reportPath), { recursive: true });
      fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), 'utf8');
    }

    console.log(`Created: ${args.output}`);
    console.log(`Slides: ${report.totals.slides}`);
    console.log(`Native text: ${report.totals.text}, shapes: ${report.totals.shapes}, images: ${report.totals.images}, SVG: ${report.totals.svg}, canvas fallbacks: ${report.totals.canvas}`);
    if (report.totals.unsupported || report.totals.failedImages) {
      console.log(`Review recommended: unsupported=${report.totals.unsupported}, failedImages=${report.totals.failedImages}`);
    }
  } finally {
    await browser.close();
    await new Promise((resolve) => server.close(resolve));
  }
}

main().catch((error) => {
  console.error(error?.stack || String(error));
  process.exit(1);
});
