use std::sync::{Arc, Mutex};

use vybe_runtime::heap;
use vybe_runtime::value::{ArrayBufferState, Object, ObjectKind, TypedArrayState, TypedElemKind};
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "libc:sdl",
        "allocFormat",
        Box::new(|ctx, args| {
            let Some(code) = args.first() else {
                return Value::Null;
            };
            let Some(format) = PixelFormat::read(ctx, code) else {
                return Value::Null;
            };
            let mut object = Object::new();
            object.properties.insert("format".into(), code.clone());
            object.properties.insert(
                "BitsPerPixel".into(),
                Value::I32(((code.as_i64() >> 8) & 255) as i32),
            );
            object
                .properties
                .insert("BytesPerPixel".into(), Value::I32(format.bytes as i32));
            object.properties.insert("palette".into(), Value::Null);
            for (i, channel) in ["R", "G", "B", "A"].into_iter().enumerate() {
                let mask = format.masks[i];
                object
                    .properties
                    .insert(format!("{channel}mask"), Value::I64(mask as i64));
                object.properties.insert(
                    format!("{channel}shift"),
                    Value::I32(if mask == 0 {
                        0
                    } else {
                        mask.trailing_zeros() as i32
                    }),
                );
                object.properties.insert(
                    format!("{channel}loss"),
                    Value::I32(8i32.saturating_sub(mask.count_ones() as i32)),
                );
            }
            Value::Object(heap::alloc(object))
        }),
    );
    vm.register_host_fn(
        "libc:sdl",
        "rgbaImageData",
        Box::new(|ctx, args| {
            let width = int_arg(args, 1).max(0) as usize;
            let height = int_arg(args, 2).max(0) as usize;
            if let Some(format) = args.get(3) {
                let Some(format) = PixelFormat::read(ctx, format) else {
                    return Value::Null;
                };
                let pitch = args
                    .get(4)
                    .map(|v| v.as_i32().max(0) as usize)
                    .unwrap_or(width * format.bytes);
                let Some(input) = read_span(ctx, args.first(), pitch.saturating_mul(height)) else {
                    return Value::Null;
                };
                let mut rgba = Vec::with_capacity(width.saturating_mul(height).saturating_mul(4));
                if format.bytes == 4
                    && format.palette.is_none()
                    && let Some(row_bytes) = width.checked_mul(4)
                    && pitch >= row_bytes
                {
                    let direct = format.masks == [0x000000ff, 0x0000ff00, 0x00ff0000, 0xff000000];
                    let bgra = format.masks == [0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000];
                    if direct || bgra {
                        for y in 0..height {
                            let Some(start) = y.checked_mul(pitch) else {
                                return Value::Null;
                            };
                            let Some(row) = input.get(start..start.saturating_add(row_bytes))
                            else {
                                return Value::Null;
                            };
                            if direct {
                                rgba.extend_from_slice(row);
                            } else {
                                for pixel in row.chunks_exact(4) {
                                    rgba.extend_from_slice(&[
                                        pixel[2], pixel[1], pixel[0], pixel[3],
                                    ]);
                                }
                            }
                        }
                        return image_data(rgba, width, height);
                    }
                }
                for y in 0..height {
                    for x in 0..width {
                        let Some(color) = format.decode(ctx, &input, y * pitch + x * format.bytes)
                        else {
                            return Value::Null;
                        };
                        rgba.extend_from_slice(&color);
                    }
                }
                return image_data(rgba, width, height);
            }
            let len = width.saturating_mul(height).saturating_mul(4);
            let mut bytes = bytes_from_value(Some(ctx), args.first()).unwrap_or_default();
            bytes.resize(len, 0);
            bytes.truncate(len);
            image_data(bytes, width, height)
        }),
    );

    vm.register_host_fn(
        "libc:sdl",
        "convertBlit",
        Box::new(|ctx, args| {
            convert_blit(ctx, args)
                .map(|_| Value::I32(0))
                .unwrap_or(Value::I32(-1))
        }),
    );
    vm.register_host_fn(
        "libc:sdl",
        "setPaletteColors",
        Box::new(|ctx, args| {
            set_palette_colors(ctx, args)
                .map(|_| Value::I32(0))
                .unwrap_or(Value::I32(-1))
        }),
    );

    vm.register_host_fn(
        "libc:sdl",
        "palettedImageData",
        Box::new(|ctx, args| {
            let width = int_arg(args, 2).max(0) as usize;
            let height = int_arg(args, 3).max(0) as usize;
            let count = width.saturating_mul(height);
            let mut rgba = vec![0u8; count.saturating_mul(4)];
            let palette = palette_source(Some(ctx), args.get(1));
            let first = palette_first(Some(ctx), args.get(1));

            for i in 0..count {
                let idx = read_u8(Some(ctx), args.first(), i) as usize;
                let palette_idx = idx.saturating_sub(first);
                let color = indexed_value(Some(ctx), palette.as_ref(), palette_idx);
                let (r, g, b) = rgb_from_value(Some(ctx), color.as_ref());
                let base = i * 4;
                rgba[base] = r;
                rgba[base + 1] = g;
                rgba[base + 2] = b;
                rgba[base + 3] = 255;
            }

            image_data(rgba, width, height)
        }),
    );
}

fn field(ctx: &HostContext<'_>, value: &Value, key: &str) -> Option<Value> {
    let Value::Object(obj) = value else {
        return None;
    };
    object_field(Some(ctx), &obj.lock().unwrap(), key)
}

fn set_palette_colors(ctx: &HostContext<'_>, args: &[Value]) -> Option<()> {
    let Value::Object(palette) = args.first()? else {
        return None;
    };
    let first = usize::try_from(args.get(2)?.as_i64()).ok()?;
    let count = usize::try_from(args.get(3)?.as_i64()).ok()?;
    if first.checked_add(count)? > 256 {
        return None;
    }
    let source = args.get(1)?;
    let packed = matches!(source, Value::Object(object)
        if matches!(&object.lock().unwrap().kind, ObjectKind::TypedArray(_)));
    let linear = if packed {
        Some(read_span(ctx, Some(source), count * 4)?)
    } else {
        None
    };
    // SDL copies colors: the caller may mutate or release the input array.
    let mut incoming = Vec::with_capacity(count);
    for index in 0..count {
        let color = if let Some(bytes) = &linear {
            bytes[index * 4..index * 4 + 4].try_into().ok()?
        } else {
            let value = indexed_value(Some(ctx), Some(source), index)?;
            let (r, g, b) = rgb_from_value(Some(ctx), Some(&value));
            [
                r,
                g,
                b,
                field(ctx, &value, "a")
                    .map(|v| v.as_i32() as u8)
                    .unwrap_or(255),
            ]
        };
        let mut copy = Object::new();
        for (name, byte) in ["r", "g", "b", "a"].into_iter().zip(color) {
            copy.properties.insert(name.into(), Value::I32(byte as i32));
        }
        incoming.push(Value::Object(heap::alloc(copy)));
    }
    let mut colors = (0..256)
        .map(|index| {
            let existing = palette_source(Some(ctx), args.first());
            indexed_value(Some(ctx), existing.as_ref(), index).unwrap_or(Value::I32(0))
        })
        .collect::<Vec<_>>();
    colors[first..first + count].clone_from_slice(&incoming);
    let mut array = Object::new();
    array.kind = ObjectKind::Array(colors);
    let mut palette = palette.lock().unwrap();
    palette
        .properties
        .insert("colors".into(), Value::Object(heap::alloc(array)));
    palette
        .properties
        .insert("__sdl_first".into(), Value::I32(0));
    palette
        .properties
        .insert("__sdl_count".into(), Value::I32(256));
    Some(())
}

fn read_span(ctx: &HostContext<'_>, value: Option<&Value>, len: usize) -> Option<Vec<u8>> {
    let value = value?;
    if matches!(value, Value::I32(_) | Value::I64(_) | Value::F64(_)) {
        let start = usize::try_from(value.as_i64()).ok()?;
        let end = start.checked_add(len)?;
        return ctx.with_linear_memory(|memory| memory.get(start..end).map(<[u8]>::to_vec))?;
    }
    let mut bytes = bytes_from_value(Some(ctx), Some(value))?;
    if bytes.len() < len {
        return None;
    }
    bytes.truncate(len);
    Some(bytes)
}

fn write_span(ctx: &mut HostContext<'_>, value: &Value, bytes: &[u8]) -> Option<()> {
    if matches!(value, Value::I32(_) | Value::I64(_) | Value::F64(_)) {
        let start = usize::try_from(value.as_i64()).ok()?;
        let end = start.checked_add(bytes.len())?;
        return ctx.with_linear_memory_mut(|memory| {
            memory
                .get_mut(start..end)
                .map(|target| target.copy_from_slice(bytes))
        })?;
    }
    if let Some((base, offset)) = carray_view(Some(ctx), value) {
        let mut tail = bytes_from_value(Some(ctx), Some(&base))?;
        let end = offset.checked_add(bytes.len())?;
        tail.get_mut(offset..end)?.copy_from_slice(bytes);
        return write_span(ctx, &base, &tail);
    }
    let Value::Object(object) = value else {
        return None;
    };
    let mut object = object.lock().unwrap();
    match &mut object.kind {
        ObjectKind::TypedArray(array)
            if matches!(
                array.elem,
                TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped
            ) =>
        {
            if bytes.len() > array.length {
                return None;
            }
            let start = array.byte_offset;
            let end = start.checked_add(bytes.len())?;
            array
                .buffer
                .lock()
                .unwrap()
                .get_mut(start..end)?
                .copy_from_slice(bytes);
        }
        ObjectKind::Array(items) => {
            if bytes.len() > items.len() {
                return None;
            }
            for (item, byte) in items.iter_mut().zip(bytes) {
                *item = Value::I32(i32::from(*byte));
            }
        }
        _ => return None,
    }
    Some(())
}

/// SDL packed pixels are little-endian guest memory, not WHATWG RGBA bytes.
/// Keeping the conversion here makes both texture presentation and surface
/// conversion use the same format interpretation.
struct PixelFormat {
    bytes: usize,
    masks: [u32; 4],
    palette: Option<Value>,
}

impl PixelFormat {
    fn read(ctx: &HostContext<'_>, value: &Value) -> Option<Self> {
        if matches!(value, Value::Object(_)) {
            let bits = field(ctx, value, "BitsPerPixel")?.as_i32();
            let bytes = ((bits + 7) / 8) as usize;
            let mut masks = [0; 4];
            for (i, name) in ["Rmask", "Gmask", "Bmask", "Amask"].iter().enumerate() {
                masks[i] = field(ctx, value, name)?.as_i64() as u32;
            }
            if masks[..3] == [0, 0, 0] {
                let defaults = match bits {
                    8 => [0, 0, 0],
                    15 => [0x7c00, 0x3e0, 0x1f],
                    16 => [0xf800, 0x7e0, 0x1f],
                    24 | 32 => [0xff0000, 0xff00, 0xff],
                    _ => return None,
                };
                masks[..3].copy_from_slice(&defaults);
            }
            return (1..=4).contains(&bytes).then(|| Self {
                bytes,
                masks,
                palette: if bits == 8 {
                    field(ctx, value, "palette")
                } else {
                    None
                },
            });
        }
        let code = value.as_i64() as u32;
        let kind = (code >> 24) & 15;
        let order = (code >> 20) & 15;
        let layout = (code >> 16) & 15;
        let bytes = (code & 255) as usize;
        if !(1..=4).contains(&bytes) {
            return None;
        }
        let mut masks = [0u32; 4];
        if kind == 7 {
            let channels: &[usize] = match order {
                1 => &[0, 1, 2],
                2 => &[0, 1, 2, 3],
                3 => &[3, 0, 1, 2],
                4 => &[2, 1, 0],
                5 => &[2, 1, 0, 3],
                6 => &[3, 2, 1, 0],
                _ => return None,
            };
            if channels.len() != bytes {
                return None;
            }
            for (i, &channel) in channels.iter().enumerate() {
                masks[channel] = 255 << (8 * i);
            }
        } else if (4..=6).contains(&kind) {
            let widths: &[u32] = match layout {
                1 => &[3, 3, 2],
                2 => &[4, 4, 4, 4],
                3 => &[1, 5, 5, 5],
                4 => &[5, 5, 5, 1],
                5 => &[5, 6, 5],
                6 => &[8, 8, 8, 8],
                7 => &[2, 10, 10, 10],
                8 => &[10, 10, 10, 2],
                _ => return None,
            };
            let channels = match order {
                1 => [4, 0, 1, 2],
                2 => [0, 1, 2, 4],
                3 => [3, 0, 1, 2],
                4 => [0, 1, 2, 3],
                5 => [4, 2, 1, 0],
                6 => [2, 1, 0, 4],
                7 => [3, 2, 1, 0],
                8 => [2, 1, 0, 3],
                _ => return None,
            };
            let channels: Vec<_> = channels
                .into_iter()
                .filter(|&c| widths.len() == 4 || c != 4)
                .collect();
            if channels.len() != widths.len() {
                return None;
            }
            let mut shift = widths.iter().sum::<u32>();
            for (&channel, &width) in channels.iter().zip(widths) {
                shift -= width;
                if channel < 4 {
                    masks[channel] = ((1u32 << width) - 1) << shift;
                }
            }
        } else {
            return None;
        }
        Some(Self {
            bytes,
            masks,
            palette: None,
        })
    }

    fn decode(&self, ctx: &HostContext<'_>, bytes: &[u8], offset: usize) -> Option<[u8; 4]> {
        let pixel = bytes.get(offset..offset.checked_add(self.bytes)?)?;
        if let Some(palette) = &self.palette {
            let first = palette_first(Some(ctx), Some(palette));
            let index = usize::from(pixel[0]).checked_sub(first)?;
            let colors = palette_source(Some(ctx), Some(palette))?;
            if matches!(colors, Value::I32(_) | Value::I64(_) | Value::F64(_)) {
                let start = usize::try_from(colors.as_i64())
                    .ok()?
                    .checked_add(index * 4)?;
                return ctx
                    .memory
                    .as_deref()?
                    .get(start..start + 4)?
                    .try_into()
                    .ok();
            }
            let color = indexed_value(Some(ctx), Some(&colors), index)?;
            let (r, g, b) = rgb_from_value(Some(ctx), Some(&color));
            let a = field(ctx, &color, "a")
                .map(|a| a.as_i32() as u8)
                .unwrap_or(255);
            return Some([r, g, b, a]);
        }
        let mut packed = [0u8; 4];
        packed[..self.bytes].copy_from_slice(pixel);
        let packed = u32::from_le_bytes(packed);
        let mut color = [0, 0, 0, 255];
        for (i, &mask) in self.masks.iter().enumerate() {
            if mask != 0 {
                let shift = mask.trailing_zeros();
                color[i] =
                    ((((packed & mask) >> shift) as u64 * 255) / u64::from(mask >> shift)) as u8;
            }
        }
        Some(color)
    }

    fn encode(&self, color: [u8; 4], out: &mut [u8]) -> Option<()> {
        if self.palette.is_some() {
            return None;
        }
        let mut packed = 0u32;
        for (component, &mask) in color.into_iter().zip(&self.masks) {
            if mask != 0 {
                let shift = mask.trailing_zeros();
                packed |= ((((u64::from(component) * u64::from(mask >> shift) + 127) / 255)
                    as u32)
                    << shift)
                    & mask;
            }
        }
        out.copy_from_slice(&packed.to_le_bytes()[..self.bytes]);
        Some(())
    }
}

fn convert_blit(ctx: &mut HostContext<'_>, args: &[Value]) -> Option<()> {
    let src = args.first()?;
    let dst = args.get(2)?;
    let sf = PixelFormat::read(ctx, &field(ctx, src, "format")?)?;
    let df = PixelFormat::read(ctx, &field(ctx, dst, "format")?)?;
    let palette = match sf.palette.as_ref() {
        Some(palette) => Some(palette_lookup(ctx, palette)?),
        None => None,
    };
    let dimension = |obj: &Value, key| usize::try_from(field(ctx, obj, key)?.as_i64()).ok();
    let (sw, sh) = (dimension(src, "w")?, dimension(src, "h")?);
    let (dw, dh) = (dimension(dst, "w")?, dimension(dst, "h")?);
    let (sp, dp) = (dimension(src, "pitch")?, dimension(dst, "pitch")?);
    let rect = |index, key, default| {
        args.get(index)
            .and_then(|r| field(ctx, r, key))
            .map(|v| usize::try_from(v.as_i64()).ok())
            .unwrap_or(Some(default))
    };
    let (sx, sy, w, h) = (
        rect(1, "x", 0)?,
        rect(1, "y", 0)?,
        rect(1, "w", sw)?,
        rect(1, "h", sh)?,
    );
    let (dx, dy) = (rect(3, "x", 0)?, rect(3, "y", 0)?);
    let scaled = args.get(4).is_some_and(|v| v.as_i32() != 0);
    let (tw, th) = if scaled {
        (rect(3, "w", dw)?, rect(3, "h", dh)?)
    } else {
        (w, h)
    };
    if sx.checked_add(w)? > sw
        || sy.checked_add(h)? > sh
        || dx.checked_add(tw)? > dw
        || dy.checked_add(th)? > dh
    {
        return None;
    }
    if tw == 0 || th == 0 || w == 0 || h == 0 {
        return Some(());
    }
    // LowerBlit receives a pre-clipped region. A surface can borrow a locked
    // texture's storage/pitch without matching its full advertised extent.
    // Validate and copy only the prefix containing the rows we access.
    let source_row_end = sx.checked_add(w)?.checked_mul(sf.bytes)?;
    let dest_row_end = dx.checked_add(tw)?.checked_mul(df.bytes)?;
    if source_row_end > sp || dest_row_end > dp {
        return None;
    }
    let input_len = sy
        .checked_add(h - 1)?
        .checked_mul(sp)?
        .checked_add(source_row_end)?;
    let output_len = dy
        .checked_add(th - 1)?
        .checked_mul(dp)?
        .checked_add(dest_row_end)?;
    let input = read_span(ctx, Some(&field(ctx, src, "pixels")?), input_len)?;
    let destination = field(ctx, dst, "pixels")?;
    let mut output = read_span(ctx, Some(&destination), output_len)?;
    for y in 0..th {
        for x in 0..tw {
            let source_y = sy + y * h / th;
            let source_x = sx + x * w / tw;
            let source_offset = source_y * sp + source_x * sf.bytes;
            let color = if let Some((first, colors)) = &palette {
                let index = usize::from(*input.get(source_offset)?).checked_sub(*first)?;
                *colors.get(index)?
            } else {
                sf.decode(ctx, &input, source_offset)?
            };
            let start = (dy + y) * dp + (dx + x) * df.bytes;
            df.encode(color, &mut output[start..start + df.bytes])?;
        }
    }
    write_span(ctx, &destination, &output)
}

fn palette_lookup(ctx: &HostContext<'_>, palette: &Value) -> Option<(usize, Vec<[u8; 4]>)> {
    let first = palette_first(Some(ctx), Some(palette));
    let count = field(ctx, palette, "ncolors")
        .map(|value| usize::try_from(value.as_i64()).ok())
        .unwrap_or(Some(256))?
        .min(256);
    let source = palette_source(Some(ctx), Some(palette))?;
    if matches!(source, Value::I32(_) | Value::I64(_) | Value::F64(_)) {
        let bytes = read_span(ctx, Some(&source), count.checked_mul(4)?)?;
        return Some((
            first,
            bytes
                .chunks_exact(4)
                .map(|part| part.try_into().unwrap())
                .collect(),
        ));
    }
    let mut colors = Vec::with_capacity(count);
    for index in 0..count {
        let value = indexed_value(Some(ctx), Some(&source), index)?;
        let (r, g, b) = rgb_from_value(Some(ctx), Some(&value));
        let a = field(ctx, &value, "a")
            .map(|value| value.as_i32() as u8)
            .unwrap_or(255);
        colors.push([r, g, b, a]);
    }
    Some((first, colors))
}

fn int_arg(args: &[Value], idx: usize) -> i32 {
    args.get(idx).map(|v| v.as_i32()).unwrap_or(0)
}

fn object_field(ctx: Option<&HostContext<'_>>, o: &Object, key: &str) -> Option<Value> {
    if let Some(value) = o.properties.get(key) {
        return Some(value.clone());
    }
    let ctx = ctx?;
    for (idx, (name, _enumerable)) in ctx.declared_fields(o.type_id).into_iter().enumerate() {
        if name == key {
            return o.fields.get(idx).cloned();
        }
    }
    None
}

fn carray_view(ctx: Option<&HostContext<'_>>, value: &Value) -> Option<(Value, usize)> {
    let Value::Object(obj) = value else {
        return None;
    };
    let o = obj.lock().unwrap();
    let kind = object_field(ctx, &o, "__ref_kind")?;
    if format!("{}", kind) != "carray" {
        return None;
    }
    let base = object_field(ctx, &o, "__base")?;
    let idx = object_field(ctx, &o, "__idx")
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0);
    Some((base, idx))
}

fn indexed_value(
    ctx: Option<&HostContext<'_>>,
    value: Option<&Value>,
    index: usize,
) -> Option<Value> {
    let value = value?;
    if let Some((base, offset)) = carray_view(ctx, value) {
        return indexed_value(ctx, Some(&base), offset.saturating_add(index));
    }
    let Value::Object(obj) = value else {
        return None;
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(items) => items.get(index).cloned(),
        ObjectKind::TypedArray(ta) => typed_array_value(ta, index),
        _ => o.properties.get(&index.to_string()).cloned(),
    }
}

fn typed_array_value(ta: &TypedArrayState, index: usize) -> Option<Value> {
    let bpe = ta.elem.bytes_per_element();
    let bytes = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset.checked_add(index.checked_mul(bpe)?)?;
    if abs.checked_add(bpe)? > bytes.len() {
        return None;
    }
    Some(match ta.elem {
        TypedElemKind::I8 => Value::I32(bytes[abs] as i8 as i32),
        TypedElemKind::U8 | TypedElemKind::U8Clamped => Value::I32(bytes[abs] as i32),
        TypedElemKind::I16 => {
            let mut raw = [0; 2];
            raw.copy_from_slice(&bytes[abs..abs + 2]);
            Value::I32(i16::from_le_bytes(raw) as i32)
        }
        TypedElemKind::U16 => {
            let mut raw = [0; 2];
            raw.copy_from_slice(&bytes[abs..abs + 2]);
            Value::I32(u16::from_le_bytes(raw) as i32)
        }
        TypedElemKind::I32 => {
            let mut raw = [0; 4];
            raw.copy_from_slice(&bytes[abs..abs + 4]);
            Value::I32(i32::from_le_bytes(raw))
        }
        TypedElemKind::U32 => {
            let mut raw = [0; 4];
            raw.copy_from_slice(&bytes[abs..abs + 4]);
            Value::I64(u32::from_le_bytes(raw) as i64)
        }
        TypedElemKind::F32 => {
            let mut raw = [0; 4];
            raw.copy_from_slice(&bytes[abs..abs + 4]);
            Value::F64(f32::from_le_bytes(raw) as f64)
        }
        TypedElemKind::F64 => {
            let mut raw = [0; 8];
            raw.copy_from_slice(&bytes[abs..abs + 8]);
            Value::F64(f64::from_le_bytes(raw))
        }
        TypedElemKind::BigI64 => {
            let mut raw = [0; 8];
            raw.copy_from_slice(&bytes[abs..abs + 8]);
            Value::I64(i64::from_le_bytes(raw))
        }
        TypedElemKind::BigU64 => {
            let mut raw = [0; 8];
            raw.copy_from_slice(&bytes[abs..abs + 8]);
            Value::I64(u64::from_le_bytes(raw) as i64)
        }
    })
}

fn read_u8(ctx: Option<&HostContext<'_>>, value: Option<&Value>, index: usize) -> u8 {
    indexed_value(ctx, value, index)
        .map(|v| v.as_i32().clamp(0, 255) as u8)
        .unwrap_or(0)
}

fn bytes_from_value(ctx: Option<&HostContext<'_>>, value: Option<&Value>) -> Option<Vec<u8>> {
    let value = value?;
    if let Some((base, offset)) = carray_view(ctx, value) {
        let mut bytes = bytes_from_value(ctx, Some(&base))?;
        let start = offset.min(bytes.len());
        bytes.drain(..start);
        return Some(bytes);
    }
    let Value::Object(obj) = value else {
        return None;
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(items) => Some(
            items
                .iter()
                .map(|v| v.as_i32().clamp(0, 255) as u8)
                .collect(),
        ),
        ObjectKind::TypedArray(ta)
            if matches!(
                ta.elem,
                TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped
            ) =>
        {
            let bytes = ta.buffer.lock().unwrap();
            let start = ta.byte_offset.min(bytes.len());
            let end = start.saturating_add(ta.length).min(bytes.len());
            Some(bytes[start..end].to_vec())
        }
        ObjectKind::TypedArray(ta) => Some(
            (0..ta.length)
                .map(|i| {
                    typed_array_value(ta, i)
                        .map(|v| v.as_i32().clamp(0, 255) as u8)
                        .unwrap_or(0)
                })
                .collect(),
        ),
        _ => None,
    }
}

fn palette_source(ctx: Option<&HostContext<'_>>, palette: Option<&Value>) -> Option<Value> {
    let Some(Value::Object(obj)) = palette else {
        return palette.cloned();
    };
    let o = obj.lock().unwrap();
    object_field(ctx, &o, "colors")
        .or_else(|| object_field(ctx, &o, "__sdl_colors"))
        .or_else(|| palette.cloned())
}

fn palette_first(ctx: Option<&HostContext<'_>>, palette: Option<&Value>) -> usize {
    let Some(Value::Object(obj)) = palette else {
        return 0;
    };
    let o = obj.lock().unwrap();
    object_field(ctx, &o, "__sdl_first")
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0)
}

fn rgb_from_value(ctx: Option<&HostContext<'_>>, value: Option<&Value>) -> (u8, u8, u8) {
    let Some(value) = value else {
        return (0, 0, 0);
    };
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        if let (Some(r), Some(g), Some(b)) = (
            object_field(ctx, &o, "r"),
            object_field(ctx, &o, "g"),
            object_field(ctx, &o, "b"),
        ) {
            return (
                r.as_i32().clamp(0, 255) as u8,
                g.as_i32().clamp(0, 255) as u8,
                b.as_i32().clamp(0, 255) as u8,
            );
        }
    }
    let packed = value.as_i32() as u32;
    (
        ((packed >> 16) & 0xff) as u8,
        ((packed >> 8) & 0xff) as u8,
        (packed & 0xff) as u8,
    )
}

fn image_data(bytes: Vec<u8>, width: usize, height: usize) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("ImageData")));
    obj.properties
        .insert("data".into(), typed_u8_clamped_array(bytes));
    obj.properties
        .insert("width".into(), Value::I32(width as i32));
    obj.properties
        .insert("height".into(), Value::I32(height as i32));
    Value::Object(heap::alloc(obj))
}

fn typed_u8_clamped_array(bytes: Vec<u8>) -> Value {
    typed_byte_array(bytes, TypedElemKind::U8Clamped, "Uint8ClampedArray")
}

fn typed_byte_array(bytes: Vec<u8>, elem: TypedElemKind, name: &str) -> Value {
    let len = bytes.len();
    let shared = Arc::new(Mutex::new(bytes));
    let ab_state = ArrayBufferState {
        bytes: shared.clone(),
        max_byte_length: len,
        resizable: false,
        detached: false,
        shared: false,
    };
    let mut ab = Object::new();
    ab.kind = ObjectKind::ArrayBuffer(ab_state);
    ab.properties
        .insert("byteLength".into(), Value::I32(len as i32));
    ab.properties
        .insert("maxByteLength".into(), Value::I32(len as i32));
    let buffer_obj = heap::alloc(ab);

    let ta_state = TypedArrayState {
        elem,
        buffer: shared,
        buffer_obj: buffer_obj.clone(),
        byte_offset: 0,
        length: len,
    };
    let mut ta = Object::new();
    ta.kind = ObjectKind::TypedArray(ta_state);
    ta.properties
        .insert("buffer".into(), Value::Object(buffer_obj));
    ta.properties
        .insert("length".into(), Value::I32(len as i32));
    ta.properties
        .insert("byteLength".into(), Value::I32(len as i32));
    ta.properties.insert("byteOffset".into(), Value::I32(0));
    ta.properties
        .insert("BYTES_PER_ELEMENT".into(), Value::I32(1));
    ta.properties
        .insert("__type".into(), Value::String(Arc::from(name)));
    ta.properties
        .insert("tostringtag".into(), Value::String(Arc::from(name)));
    Value::Object(heap::alloc(ta))
}
