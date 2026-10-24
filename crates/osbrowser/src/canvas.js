// webcanvas's `DrawCmd` transcript, replayed onto the page's real 2D context.
// Each command is `[method, ...args]`, recorded by osbrowser's canvas.
window.osbrowserCanvas2D = (() => {
  const css = c => `rgba(${c[0]},${c[1]},${c[2]},${c[3] / 255})`;
  const scratch = document.createElement('canvas').getContext('2d');

  function gradient(ctx, g) {
    const out = g.linear ? ctx.createLinearGradient(...g.linear)
      : g.radial ? ctx.createRadialGradient(...g.radial)
      : ctx.createConicGradient(...g.conic);
    for (const [offset, color] of g.stops) out.addColorStop(offset, css(color));
    return out;
  }

  function path(ops, target = new Path2D()) {
    for (const [method, ...args] of ops) target[method](...args);
    return target;
  }

  function surface(info, pixels) {
    const canvas = document.createElement('canvas');
    canvas.width = info.width;
    canvas.height = info.height;
    if (info.width && info.height) {
      canvas.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(pixels), info.width, info.height), 0, 0);
    }
    return canvas;
  }

  function draw(ctx, canvas, commands) {
    for (const [method, ...a] of commands) {
      switch (method) {
        case 'noop': break;
        case 'set': ctx[a[0]] = a[1]; break;
        case 'fillColor': ctx.fillStyle = css(a[0]); break;
        case 'strokeColor': ctx.strokeStyle = css(a[0]); break;
        case 'fillGradient': ctx.fillStyle = gradient(ctx, a[0]); break;
        case 'strokeGradient': ctx.strokeStyle = gradient(ctx, a[0]); break;
        case 'shadow':
          ctx.shadowColor = css(a[0]); ctx.shadowBlur = a[1];
          ctx.shadowOffsetX = a[2]; ctx.shadowOffsetY = a[3];
          break;
        case 'setTransform': ctx.setTransform(...a[0]); break;
        case 'reset':
          if (ctx.reset) ctx.reset(); else canvas.width = canvas.width;
          break;
        case 'resetClip': ctx.resetClip?.(); break;
        case 'drawFocusIfNeeded': break;
        default: ctx[method](...a);
      }
    }
  }

  function image(ctx, command, pixels) {
    const [method, info, ...a] = command;
    switch (method) {
      case 'putImageData':
        if (info.width && info.height) {
          ctx.putImageData(new ImageData(new Uint8ClampedArray(pixels), info.width, info.height), a[0], a[1]);
        }
        break;
      case 'drawImage': ctx.drawImage(surface(info, pixels), a[0], a[1], a[2], a[3]); break;
      case 'drawImageRect': ctx.drawImage(surface(info, pixels), ...a); break;
      case 'fillPattern':
      case 'strokePattern': {
        const pattern = ctx.createPattern(surface(info, pixels), a[0]);
        if (pattern) pattern.setTransform(new DOMMatrix(a[1]));
        ctx[method === 'fillPattern' ? 'fillStyle' : 'strokeStyle'] = pattern;
        break;
      }
    }
  }

  const base64 = bytes => {
    let binary = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(binary);
  };

  function query(ctx, canvas, question, a) {
    switch (question) {
      case 'measureText': {
        const m = ctx.measureText(a[0]);
        return {
          width: m.width,
          actual_bounding_box_left: m.actualBoundingBoxLeft || 0,
          actual_bounding_box_right: m.actualBoundingBoxRight || 0,
          font_bounding_box_ascent: m.fontBoundingBoxAscent || 0,
          font_bounding_box_descent: m.fontBoundingBoxDescent || 0,
          actual_bounding_box_ascent: m.actualBoundingBoxAscent || 0,
          actual_bounding_box_descent: m.actualBoundingBoxDescent || 0,
          em_height_ascent: m.emHeightAscent || 0,
          em_height_descent: m.emHeightDescent || 0,
          hanging_baseline: m.hangingBaseline || 0,
          alphabetic_baseline: m.alphabeticBaseline || 0,
          ideographic_baseline: m.ideographicBaseline || 0,
        };
      }
      case 'isPointInPath': return ctx.isPointInPath(a[0], a[1], a[2]);
      case 'isPointInStroke': return ctx.isPointInStroke(a[0], a[1]);
      case 'isPointInPath2D': return ctx.isPointInPath(path(a[0]), a[1], a[2], a[3]);
      case 'isPointInStroke2D': return ctx.isPointInStroke(path(a[0]), a[1], a[2]);
      case 'getImageData': {
        const data = ctx.getImageData(a[0], a[1], a[2], a[3]);
        return { width: data.width, height: data.height, data: base64(data.data) };
      }
      case 'toBlob': {
        const url = canvas.toDataURL(a[0], a[1] ?? undefined);
        const comma = url.indexOf(',');
        return comma < 0 ? null : url.slice(comma + 1);
      }
      case 'isContextLost': return Boolean(ctx.isContextLost?.());
      default: throw new Error(`unsupported canvas query: ${question}`);
    }
  }

  // The page's own CSS parser, read back through the canvas's serialisation.
  // Two sentinels, because an invalid value leaves the previous one in place
  // and a single sentinel cannot tell "invalid" from "the sentinel's colour".
  function parseColor(value) {
    const read = sentinel => { scratch.fillStyle = sentinel; scratch.fillStyle = value; return scratch.fillStyle; };
    const first = read('#000001');
    if (first === '#000001' && read('#000002') === '#000002') return null;
    if (first[0] === '#') {
      return [1, 3, 5].map(i => parseInt(first.slice(i, i + 2), 16)).concat(255);
    }
    const parts = first.match(/[\d.]+/g)?.map(Number) || [];
    if (parts.length < 3) return null;
    return [parts[0], parts[1], parts[2], Math.round((parts[3] ?? 1) * 255)];
  }

  function parseFont(value) {
    const read = sentinel => { scratch.font = sentinel; scratch.font = value; return scratch.font; };
    const first = read('1px a');
    if (first === '1px a' && read('2px b') === '2px b') return null;
    const match = first.match(/^(.*?)(\d+(?:\.\d+)?)px(?:\/\S+)?\s+(.+)$/);
    if (!match) return null;
    const words = match[1].toLowerCase().split(/\s+/);
    const bold = words.some(w => w === 'bold' || w === 'bolder' || Number(w) >= 600);
    return { family: match[3], size: Number(match[2]), bold, italic: words.includes('italic') || words.includes('oblique') };
  }

  return { draw, image, query, parseColor, parseFont };
})();
