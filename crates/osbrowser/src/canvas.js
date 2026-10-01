// Canvas operations run in the browser's own CanvasRenderingContext2D.
window.vybeCanvas = (() => {
  const tagged = (name, value) => value === undefined ? name : { [name]: value };
  const variant = value => typeof value === 'string'
    ? [value, null]
    : Object.entries(value)[0];
  const prop = name => name[0].toLowerCase() + name.slice(1);
  const rgba = values => `rgba(${values[0]},${values[1]},${values[2]},${values[3] / 255})`;

  function targetOf(name, doc, nodes) {
    if (name.startsWith('n') && /^n\d+$/.test(name)) return nodes.get(Number(name.slice(1)));
    return doc.getElementById(name) || doc.querySelector(`[name="${CSS.escape(name)}"]`);
  }

  function path(def, target) {
    const out = target || new Path2D();
    for (const item of def.ops || []) {
      const [name, args] = variant(item);
      switch (name) {
        case 'ClosePath': out.closePath(); break;
        case 'MoveTo': out.moveTo(...args); break;
        case 'LineTo': out.lineTo(...args); break;
        case 'QuadraticCurveTo': out.quadraticCurveTo(args.cx, args.cy, args.x, args.y); break;
        case 'BezierCurveTo':
          out.bezierCurveTo(args.cx1, args.cy1, args.cx2, args.cy2, args.x, args.y); break;
        case 'ArcTo': out.arcTo(args.x1, args.y1, args.x2, args.y2, args.radius); break;
        case 'Rect': out.rect(args.x, args.y, args.w, args.h); break;
        case 'RoundRect': out.roundRect(args.x, args.y, args.w, args.h, args.radii); break;
        case 'Arc': out.arc(args.x, args.y, args.r, args.start, args.end, args.ccw); break;
        case 'Ellipse':
          out.ellipse(args.x, args.y, args.rx, args.ry, args.rotation,
            args.start, args.end, args.ccw); break;
        default: throw new Error(`unsupported path operation: ${name}`);
      }
    }
    return out;
  }

  function image(pixels, width, height) {
    const canvas = document.createElement('canvas');
    canvas.width = width; canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.putImageData(new ImageData(Uint8ClampedArray.from(pixels), width, height), 0, 0);
    return canvas;
  }

  function gradient(ctx, def) {
    const [kind, a] = variant(def.kind);
    let out;
    if (kind === 'Linear') out = ctx.createLinearGradient(a.x0, a.y0, a.x1, a.y1);
    else if (kind === 'Radial')
      out = ctx.createRadialGradient(a.x0, a.y0, a.r0, a.x1, a.y1, a.r1);
    else if (kind === 'Conic') out = ctx.createConicGradient(a.angle, a.x, a.y);
    else throw new Error(`unsupported gradient: ${kind}`);
    for (const [offset, color] of def.stops) out.addColorStop(offset, color);
    return out;
  }

  function apply(ctx, canvas, op) {
    const [name, a] = variant(op);
    const set = (field, value) => { ctx[field] = value; };
    switch (name) {
      case 'Save': ctx.save(); break;
      case 'Restore': ctx.restore(); break;
      case 'SetFillStyle': set('fillStyle', rgba(a)); break;
      case 'SetStrokeStyle': set('strokeStyle', rgba(a)); break;
      case 'SetFont':
        set('font', `${a.italic ? 'italic ' : ''}${a.bold ? 'bold ' : ''}${a.size}px ${a.family}`);
        break;
      case 'SetLineDash': ctx.setLineDash(a); break;
      case 'SetFillStyleCss': set('fillStyle', a); break;
      case 'SetStrokeStyleCss': set('strokeStyle', a); break;
      case 'SetFontCss': set('font', a); break;
      case 'SetFilter': set('filter', a); break;
      case 'SetGlobalCompositeOperation': set('globalCompositeOperation', a); break;
      case 'SetImageSmoothingQuality': set('imageSmoothingQuality', a); break;
      case 'SetFillGradient': set('fillStyle', gradient(ctx, a)); break;
      case 'SetStrokeGradient': set('strokeStyle', gradient(ctx, a)); break;
      case 'SetFillPattern':
      case 'SetStrokePattern':
        set(name === 'SetFillPattern' ? 'fillStyle' : 'strokeStyle',
          ctx.createPattern(image(a.pixels, a.width, a.height), a.repetition));
        break;
      case 'Transform': ctx.transform(...a); break;
      case 'ResetTransform': ctx.resetTransform(); break;
      case 'Translate': ctx.translate(...a); break;
      case 'Scale': ctx.scale(...a); break;
      case 'Rotate': ctx.rotate(a); break;
      case 'Arc': ctx.arc(a[0], a[1], a[2], a[3], a[4], a[5]); break;
      case 'Ellipse': ctx.ellipse(a[0], a[1], a[2], a[3], 0, 0, Math.PI * 2); break;
      case 'EllipseFull':
        ctx.ellipse(a.x, a.y, a.rx, a.ry, a.rotation, a.start, a.end, a.ccw); break;
      case 'RoundRect': ctx.roundRect(a.x, a.y, a.w, a.h, a.radii); break;
      case 'FillWithRule': ctx.fill(a); break;
      case 'ClipWithRule': ctx.clip(a); break;
      case 'FillTextMaxWidth': ctx.fillText(...a); break;
      case 'StrokeTextMaxWidth': ctx.strokeText(...a); break;
      case 'FillPath': ctx.fill(path(a[0]), a[1]); break;
      case 'StrokePath': ctx.stroke(path(a)); break;
      case 'ClipPath': ctx.clip(path(a[0]), a[1]); break;
      case 'AppendPath': path(a, ctx); break;
      case 'PutImageData':
        ctx.putImageData(new ImageData(Uint8ClampedArray.from(a.pixels), a.width, a.height),
          a.dx, a.dy); break;
      case 'PutImageDataDirty':
        ctx.putImageData(new ImageData(Uint8ClampedArray.from(a.pixels), a.width, a.height),
          a.dx, a.dy, a.dirty_x, a.dirty_y, a.dirty_w, a.dirty_h); break;
      case 'DrawImageRgba':
        ctx.drawImage(image(a.pixels, a.width, a.height), a.dx, a.dy, a.dw, a.dh); break;
      case 'Reset':
        if (ctx.reset) ctx.reset(); else canvas.width = canvas.width;
        break;
      case 'DrawFocusIfNeeded':
        if (a && document.activeElement === canvas) ctx.drawFocusIfNeeded(canvas);
        break;
      default: {
        const property = {
          SetLineWidth: 'lineWidth', SetLineCap: 'lineCap', SetLineJoin: 'lineJoin',
          SetGlobalAlpha: 'globalAlpha', SetImageSmoothing: 'imageSmoothingEnabled',
          SetMiterLimit: 'miterLimit', SetLineDashOffset: 'lineDashOffset',
          SetTextAlign: 'textAlign', SetTextBaseline: 'textBaseline',
          SetShadowColor: 'shadowColor', SetShadowBlur: 'shadowBlur',
          SetShadowOffsetX: 'shadowOffsetX', SetShadowOffsetY: 'shadowOffsetY',
          SetDirection: 'direction', SetLetterSpacing: 'letterSpacing',
          SetWordSpacing: 'wordSpacing', SetFontKerning: 'fontKerning',
          SetFontStretch: 'fontStretch', SetFontVariantCaps: 'fontVariantCaps',
          SetTextRendering: 'textRendering', SetLang: 'lang'
        }[name];
        if (property) { set(property, a); break; }
        const method = {
          BeginPath: 'beginPath', ClosePath: 'closePath', MoveTo: 'moveTo',
          LineTo: 'lineTo', ArcTo: 'arcTo', BezierCurveTo: 'bezierCurveTo',
          QuadraticCurveTo: 'quadraticCurveTo', Rect: 'rect', Fill: 'fill',
          Stroke: 'stroke', Clip: 'clip', FillRect: 'fillRect',
          StrokeRect: 'strokeRect', ClearRect: 'clearRect',
          FillText: 'fillText', StrokeText: 'strokeText'
        }[name];
        if (!method) throw new Error(`unsupported canvas operation: ${name}`);
        ctx[method](...(Array.isArray(a) ? a : a === null ? [] : [a]));
      }
    }
    return 'None';
  }

  function query(ctx, canvas, op) {
    const [name, a] = variant(op);
    switch (name) {
      case 'MeasureText': {
        const m = ctx.measureText(a);
        const fields = ['width', 'actualBoundingBoxLeft', 'actualBoundingBoxRight',
          'actualBoundingBoxAscent', 'actualBoundingBoxDescent',
          'fontBoundingBoxAscent', 'fontBoundingBoxDescent', 'emHeightAscent',
          'emHeightDescent', 'hangingBaseline', 'alphabeticBaseline',
          'ideographicBaseline'];
        const values = {};
        for (const field of fields) {
          const snake = field.replace(/[A-Z]/g, char => `_${char.toLowerCase()}`);
          values[snake] = m[field] || 0;
        }
        return tagged('Metrics', values);
      }
      case 'GetImageData': {
        const data = ctx.getImageData(a.sx, a.sy, a.sw, a.sh);
        return tagged('Pixels', { data: Array.from(data.data), width: data.width, height: data.height });
      }
      case 'SnapshotSource': {
        const data = ctx.getImageData(a.sx, a.sy, a.sw, a.sh);
        return tagged('SourceImage', { data: Array.from(data.data),
          width: data.width, height: data.height, origin_clean: true });
      }
      case 'IsPointInPath': return tagged('Bool', ctx.isPointInPath(a.x, a.y, a.rule));
      case 'IsPointInStroke': return tagged('Bool', ctx.isPointInStroke(a.x, a.y));
      case 'GetTransform': {
        const m = ctx.getTransform();
        return tagged('Matrix', [m.a, m.b, m.c, m.d, m.e, m.f]);
      }
      case 'GetLineDash': return tagged('Floats', ctx.getLineDash());
      case 'IsContextLost': return tagged('Bool', ctx.isContextLost?.() || false);
      case 'ToDataUrl': return tagged('Text', canvas.toDataURL(a.mime, a.quality ?? undefined));
      case 'ToBlob': {
        const url = canvas.toDataURL(a.mime, a.quality ?? undefined);
        const bytes = atob(url.split(',')[1]);
        return tagged('Bytes', Array.from(bytes, char => char.charCodeAt(0)));
      }
      case 'GetContextAttributes': {
        const attrs = ctx.getContextAttributes?.() || {};
        return tagged('ContextAttributes', {
          alpha: attrs.alpha ?? true, desynchronized: attrs.desynchronized ?? false,
          color_space: attrs.colorSpace || 'srgb', color_type: attrs.colorType || 'unorm8',
          will_read_frequently: attrs.willReadFrequently ?? false
        });
      }
      case 'IsPointInPathOf': return tagged('Bool', ctx.isPointInPath(path(a.path), a.x, a.y, a.rule));
      case 'IsPointInStrokeOf': return tagged('Bool', ctx.isPointInStroke(path(a.path), a.x, a.y));
      case 'GetStringAttribute': return tagged('Text', String(ctx[prop(a)] ?? ''));
      default: throw new Error(`unsupported canvas query: ${name}`);
    }
  }

  return (operation, args, doc, nodes) => {
    const canvas = targetOf(args.target, doc, nodes);
    if (!(canvas instanceof HTMLCanvasElement)) return 'Absent';
    if (operation === 'CanvasClearAll') { canvas.width = canvas.width; return 'None'; }
    const ctx = canvas.getContext('2d');
    if (!ctx) return 'Absent';
    if (operation === 'CanvasEnsure') return 'None';
    if (operation === 'CanvasApply') return apply(ctx, canvas, args.payload);
    if (operation === 'CanvasQuery') return query(ctx, canvas, args.payload);
    throw new Error(`unsupported canvas command: ${operation}`);
  };
})();
