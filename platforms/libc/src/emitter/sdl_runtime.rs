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
                let Some(row_bytes) = width.checked_mul(format.bytes) else {
                    return Value::Null;
                };
                let Some(output_len) = width.checked_mul(height).and_then(|count| count.checked_mul(4))
                    else { return Value::Null; };
                let pitch = args
                    .get(4)
                    .map(|v| v.as_i32().max(0) as usize)
                    .unwrap_or(row_bytes);
                if height != 0 && pitch < row_bytes { return Value::Null; }
                // Only the accessed prefix is required; final-row padding is
                // not part of the pixel data and may lie outside borrowed storage.
                let input_len = if height == 0 { Some(0) } else {
                    pitch.checked_mul(height - 1).and_then(|prefix| prefix.checked_add(row_bytes))
                };
                let Some(input_len) = input_len else { return Value::Null; };
                let Some(mut input) = read_span(ctx, args.first(), input_len) else {
                    return Value::Null;
                };
                // The source is already an independent snapshot. Move it into
                // ImageData rather than allocating and copying the same bytes again.
                if direct_rgba_format(&format, width, pitch) {
                    return image_data(input, width, height);
                }
                if format.bytes == 4
                    && format.palette.is_none()
                    && let Some(row_bytes) = width.checked_mul(4)
                    && pitch >= row_bytes
                {
                    let direct = format.masks == [0x000000ff, 0x0000ff00, 0x00ff0000, 0xff000000];
                    let bgra = format.masks == [0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000];
                    if direct || bgra {
                        if output_len == 0 {
                            return image_data(Vec::new(), width, height);
                        }
                        if bgra && reuse_owned_frame_buffer(input.capacity(), output_len) {
                            if pack_rgba_rows_in_place(&mut input, width, height, pitch, true).is_none() {
                                return Value::Null;
                            }
                            return image_data(input, width, height);
                        }
                        // Padded RGBA is already a bulk row copy. Repacking it
                        // in place can be slower; oversized snapshots should
                        // also release their padding when presentation ends.
                        let Some(rgba) = compact_rgba_rows(&input, width, height, pitch, bgra)
                            else { return Value::Null; };
                        return image_data(rgba, width, height);
                    }
                }
                if let Some(bgr) = format.rgb24_order() {
                    return frame_kernels::expand_rgb24_rows(&input, width, height, pitch, bgr)
                        .map(|rgba| image_data(rgba, width, height))
                        .unwrap_or(Value::Null);
                }
                let mut rgba = Vec::new();
                if rgba.try_reserve_exact(output_len).is_err() { return Value::Null; }
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
        "bulkBlitLength",
        Box::new(|ctx, args| {
            bulk_blit_length(ctx, args)
                .map(|length| Value::F64(f64::from(length)))
                .unwrap_or(Value::I32(0))
        }),
    );
    vm.register_host_fn(
        "libc:sdl",
        "bulkBlitRowBytes",
        Box::new(|ctx, args| {
            bulk_blit_layout(ctx, args, false)
                .map(|(row, _)| Value::F64(f64::from(row)))
                .unwrap_or(Value::I32(0))
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
        Box::new(|ctx, args| paletted_image_data(ctx, args)),
    );
}

fn direct_rgba_format(format: &PixelFormat, width: usize, pitch: usize) -> bool {
    format.bytes == 4 && format.palette.is_none()
        && format.masks == [0xff, 0xff00, 0xff0000, 0xff000000]
        && width.checked_mul(4) == Some(pitch)
}

#[path = "sdl_pixel_kernels.rs"]
mod frame_kernels;

/// Bound the unused padding capacity retained by a reused snapshot.
fn reuse_owned_frame_buffer(capacity: usize, output_len: usize) -> bool {
    output_len != 0 && output_len.checked_add(output_len / 8)
        .is_some_and(|limit| capacity <= limit)
}

/// Produce a compact owned result without retaining oversized source padding.
fn compact_rgba_rows(
    input: &[u8], width: usize, rows: usize, pitch: usize, bgra: bool,
) -> Option<Vec<u8>> {
    let row_bytes = width.checked_mul(4)?;
    let output_len = row_bytes.checked_mul(rows)?;
    if rows != 0 && pitch < row_bytes { return None; }
    let source_len = if rows == 0 { 0 } else {
        (rows - 1).checked_mul(pitch)?.checked_add(row_bytes)?
    };
    if source_len > input.len() { return None; }
    let mut output = Vec::new();
    output.try_reserve_exact(output_len).ok()?;
    if row_bytes != 0 {
        for row in 0..rows {
            let start = row * pitch;
            output.extend_from_slice(&input[start..start + row_bytes]);
        }
        if bgra { frame_kernels::swap_bgra_bytes(&mut output); }
    }
    Some(output)
}

/// Convert only an independently owned snapshot. Validate every accessed row
/// before moving bytes; source rows follow their compacted destination rows.
fn pack_rgba_rows_in_place(
    bytes: &mut Vec<u8>, width: usize, rows: usize, pitch: usize, bgra: bool,
) -> Option<()> {
    let row_bytes = width.checked_mul(4)?;
    let output_len = row_bytes.checked_mul(rows)?;
    if rows != 0 && pitch < row_bytes { return None; }
    let source_len = if rows == 0 { 0 } else {
        (rows - 1).checked_mul(pitch)?.checked_add(row_bytes)?
    };
    if source_len > bytes.len() { return None; }
    if row_bytes != 0 {
        for row in 0..rows {
            let source = row * pitch;
            let destination = row * row_bytes;
            if source != destination {
                bytes.copy_within(source..source + row_bytes, destination);
            }
        }
        if bgra {
            frame_kernels::swap_bgra_bytes(&mut bytes[..output_len]);
        }
    }
    bytes.truncate(output_len);
    Some(())
}

fn expand_palette_indices(indices: &[u8], palette: &[[u8; 4]; 256], rgba: &mut [u8]) {
    for (index, pixel) in indices.iter().zip(rgba.chunks_exact_mut(4)) {
        pixel.copy_from_slice(&palette[usize::from(*index)]);
    }
}

/// Borrow ordinary byte representations once, without executing conversions
/// that could lock a nested object while its parent is locked.
fn expand_managed_palette_indices(
    ctx: &HostContext<'_>, source: &Value, offset: usize, count: usize,
    palette: &[[u8; 4]; 256], rgba: &mut [u8],
) -> Option<()> {
    if let Some((base, view_offset)) = carray_view(Some(ctx), source) {
        return expand_managed_palette_indices(ctx, &base, offset.checked_add(view_offset)?,
            count, palette, rgba);
    }
    let Value::Object(object) = source else { return None; };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(values) => {
            let values = values.get(offset..offset.checked_add(count)?)?;
            if !values.iter().all(|value| matches!(value,
                Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_)
                | Value::Bool(_) | Value::Null | Value::Undefined | Value::String(_))) {
                return None;
            }
            for (value, pixel) in values.iter().zip(rgba.chunks_exact_mut(4)) {
                pixel.copy_from_slice(&palette[value.as_i32().clamp(0, 255) as usize]);
            }
        }
        ObjectKind::TypedArray(array) if matches!(array.elem,
            TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped) => {
            if offset.checked_add(count)? > array.length { return None; }
            let start = array.byte_offset.checked_add(offset)?;
            let end = start.checked_add(count)?;
            let buffer = array.buffer.lock().unwrap();
            let indices = buffer.get(start..end)?;
            if matches!(array.elem, TypedElemKind::I8) {
                // Preserve the previous signed-element numeric conversion.
                for (index, pixel) in indices.iter().zip(rgba.chunks_exact_mut(4)) {
                    pixel.copy_from_slice(&palette[(*index as i8).max(0) as usize]);
                }
            } else {
                expand_palette_indices(indices, palette, rgba);
            }
        }
        _ => return None,
    }
    Some(())
}

fn paletted_image_data(ctx: &HostContext<'_>, args: &[Value]) -> Value {
    let width = int_arg(args, 2).max(0) as usize;
    let height = int_arg(args, 3).max(0) as usize;
    let Some(count) = width.checked_mul(height) else { return Value::Null; };
    let Some(len) = count.checked_mul(4) else { return Value::Null; };
    // Reject invalid linear spans before reserving/zeroing a possibly huge
    // framebuffer. Shared WASM memory may grow, but cannot shrink this range.
    let linear_span = if count != 0 && args.first().is_some_and(|source|
        matches!(source, Value::I32(_) | Value::I64(_) | Value::F64(_))) {
        let Some(start) = args.first().and_then(linear_memory32_address)
            .map(|address| address as usize) else { return Value::Null; };
        let Some(end) = start.checked_add(count) else { return Value::Null; };
        if ctx.with_linear_memory(|memory| end <= memory.len()) != Some(true) {
            return Value::Null;
        }
        Some((start, end))
    } else { None };
    let mut rgba = Vec::new();
    if rgba.try_reserve_exact(len).is_err() { return Value::Null; }
    rgba.resize(len, 0);
    if count == 0 { return image_data(rgba, width, height); }

    // Rebuild per call, so palette mutations are visible on the next frame.
    // Preserve the adapter's existing opaque-alpha and missing-color behavior.
    let colors = palette_source(Some(ctx), args.get(1));
    let first = palette_first(Some(ctx), args.get(1));
    let mut palette = [[0, 0, 0, 255]; 256];
    for (index, pixel) in palette.iter_mut().enumerate() {
        let color = indexed_value(Some(ctx), colors.as_ref(), index.saturating_sub(first));
        let (r, g, b) = rgb_from_value(Some(ctx), color.as_ref());
        *pixel = [r, g, b, 255];
    }

    if let Some(source) = args.first() {
        if let Some((start, end)) = linear_span {
            let copied = ctx.with_linear_memory(|memory| {
                let indices = memory.get(start..end)?;
                expand_palette_indices(indices, &palette, &mut rgba);
                Some(())
            }).flatten();
            if copied.is_none() { return Value::Null; }
        } else if expand_managed_palette_indices(ctx, source, 0, count, &palette, &mut rgba).is_none() {
            for (index, pixel) in rgba.chunks_exact_mut(4).enumerate() {
                pixel.copy_from_slice(&palette[usize::from(read_u8(Some(ctx), Some(source), index))]);
            }
        }
    } else {
        for pixel in rgba.chunks_exact_mut(4) { pixel.copy_from_slice(&palette[0]); }
    }
    image_data(rgba, width, height)
}

#[cfg(test)]
mod paletted_bulk_tests {
    use super::*;

    fn colors() -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::Array(vec![Value::I32(0x123456), Value::I32(0xabcdef)]);
        Value::Object(heap::alloc(object))
    }

    fn pixels(ctx: &HostContext<'_>, image: &Value) -> Vec<u8> {
        let data = field(ctx, image, "data").unwrap();
        bytes_from_value(Some(ctx), Some(&data)).unwrap()
    }

    #[test]
    fn managed_and_linear_pixels_match_and_output_is_independent() {
        let mut memory = [9, 0, 1, 2, 255, 9];
        let mut ctx = HostContext::empty();
        ctx.memory = Some(&mut memory);
        let palette = colors();
        let managed = typed_byte_array(vec![0, 1, 2, 255], TypedElemKind::U8, "Uint8Array");
        let image = paletted_image_data(&ctx,
            &[Value::I32(1), palette.clone(), Value::I32(4), Value::I32(1)]);
        let expected = vec![0x12, 0x34, 0x56, 255, 0xab, 0xcd, 0xef, 255,
            0, 0, 0, 255, 0, 0, 0, 255];
        assert_eq!(pixels(&ctx, &image), expected);
        assert_eq!(pixels(&ctx, &paletted_image_data(&ctx,
            &[managed, palette, Value::I32(4), Value::I32(1)])), expected);
        ctx.with_linear_memory_mut(|bytes| bytes.fill(0));
        assert_eq!(pixels(&ctx, &image), expected);
    }

    #[test]
    fn signed_elements_and_short_managed_fallback_preserve_numeric_conversion() {
        let ctx = HostContext::empty();
        let signed = typed_byte_array(vec![255, 1], TypedElemKind::I8, "Int8Array");
        let image = paletted_image_data(&ctx,
            &[signed, colors(), Value::I32(2), Value::I32(1)]);
        assert_eq!(pixels(&ctx, &image), [0x12, 0x34, 0x56, 255, 0xab, 0xcd, 0xef, 255]);
        let mut short = Object::new();
        short.kind = ObjectKind::Array(vec![Value::I32(1)]);
        let image = paletted_image_data(&ctx,
            &[Value::Object(heap::alloc(short)), colors(), Value::I32(2), Value::I32(1)]);
        assert_eq!(pixels(&ctx, &image), [0xab, 0xcd, 0xef, 255, 0x12, 0x34, 0x56, 255]);
    }

    #[test]
    fn invalid_linear_range_and_impossible_output_capacity_fail_without_access() {
        let mut memory = [0; 4];
        let mut ctx = HostContext::empty();
        ctx.memory = Some(&mut memory);
        for args in [
            [Value::I32(3), colors(), Value::I32(2), Value::I32(1)],
            [Value::F64(f64::NAN), colors(), Value::I32(1), Value::I32(1)],
            [Value::I32(0), colors(), Value::I32(i32::MAX), Value::I32(i32::MAX)],
        ] {
            assert!(matches!(paletted_image_data(&ctx, &args), Value::Null));
        }
        assert_eq!(ctx.with_linear_memory(|bytes| bytes.to_vec()), Some(vec![0; 4]));
    }

    #[test]
    fn direct_rgba_move_requires_exact_channels_and_contiguous_rows() {
        let mut format = PixelFormat { bytes: 4,
            masks: [0xff, 0xff00, 0xff0000, 0xff000000], palette: None };
        assert!(direct_rgba_format(&format, 3, 12));
        assert!(!direct_rgba_format(&format, 3, 16));
        format.masks = [0xff0000, 0xff00, 0xff, 0xff000000];
        assert!(!direct_rgba_format(&format, 3, 12));
        format.masks = [0xff, 0xff00, 0xff0000, 0];
        assert!(!direct_rgba_format(&format, 3, 12));
    }
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
    let limit = field(ctx, args.first()?, "ncolors")
        .map(|value| usize::try_from(value.as_i64()).ok())
        .unwrap_or(Some(256))?.min(256);
    if first.checked_add(count)? > limit {
        return None;
    }
    if count == 0 { return Some(()); }
    let source = args.get(1)?;
    let linear = match source {
        Value::I32(_) | Value::I64(_) | Value::F64(_) => {
            let address = match source {
                // A native memory32 address is an unsigned i32 bit pattern.
                Value::I32(value) => i64::from(*value as u32),
                Value::I64(value) if *value >= 0 => *value,
                Value::F64(value) if value.is_finite() && *value >= 0.0
                    && value.trunc() == *value => *value as i64,
                _ => return None,
            };
            // SDL's colors argument is a C pointer: NULL with a nonempty
            // update is an error, not permission to read the memory-zero prefix.
            if address == 0 { return None; }
            Some(read_span(ctx, Some(&Value::I64(address)), count * 4)?)
        }
        Value::Object(object) => {
            let object = object.lock().unwrap();
            if let ObjectKind::TypedArray(array) = &object.kind {
                // SDL_Color is four consecutive bytes, even when the caller
                // reaches that storage through a wider numeric typed view.
                let len = count * 4;
                if len > array.length.checked_mul(array.elem.bytes_per_element())? {
                    return None;
                }
                let start = array.byte_offset;
                let end = start.checked_add(len)?;
                let bytes = array.buffer.lock().unwrap();
                Some(bytes.get(start..end)?.to_vec())
            } else { None }
        }
        _ => None,
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
        copy.properties.reserve(4);
        for (name, byte) in ["r", "g", "b", "a"].into_iter().zip(color) {
            copy.properties.insert(name.into(), Value::I32(byte as i32));
        }
        incoming.push(Value::Object(heap::alloc(copy)));
    }
    let existing = palette_source(Some(ctx), args.first());
    let mut colors = (0..256)
        .map(|index| {
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
    fn rgb24_order(&self) -> Option<bool> {
        if self.bytes != 3 || self.palette.is_some() { return None; }
        match self.masks {
            [0xff, 0xff00, 0xff0000, 0] => Some(false),
            [0xff0000, 0xff00, 0xff, 0] => Some(true),
            _ => None,
        }
    }

    fn byte_channels(&self) -> bool {
        if self.palette.is_some() || !(3..=4).contains(&self.bytes) {
            return false;
        }
        let mut covered = 0u32;
        for mask in self.masks {
            if mask == 0 {
                continue;
            }
            let shift = mask.trailing_zeros();
            if shift % 8 != 0 || mask >> shift != 255 || covered & mask != 0 {
                return false;
            }
            covered |= mask;
        }
        covered == if self.bytes == 4 { u32::MAX } else { 0x00ff_ffff }
    }

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
            let mut active_channels = [0usize; 4];
            let mut channel_count = 0;
            for channel in channels {
                if widths.len() == 4 || channel != 4 {
                    active_channels[channel_count] = channel;
                    channel_count += 1;
                }
            }
            if channel_count != widths.len() {
                return None;
            }
            let channels = &active_channels[..channel_count];
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

/// A representation proof only: actual copying remains standard WASM emission.
/// Source snapshot semantics agree with memory.copy for every overlap direction.
fn bulk_blit_length(ctx: &HostContext<'_>, args: &[Value]) -> Option<u32> {
    bulk_blit_layout(ctx, args, true).map(|(_, span)| span)
}

fn bulk_blit_layout(ctx: &HostContext<'_>, args: &[Value], packed: bool) -> Option<(u32, u32)> {
    let absent_rect = |value: &Value| match value {
        Value::Null | Value::Undefined => true,
        Value::I32(value) => *value == 0,
        Value::I64(value) => *value == 0,
        Value::F64(value) => *value == 0.0,
        _ => false,
    };
    if !absent_rect(args.get(1)?) || !absent_rect(args.get(3)?) {
        return None;
    }
    let src = args.first()?;
    let dst = args.get(2)?;
    let sf = PixelFormat::read(ctx, &field(ctx, src, "format")?)?;
    let df = PixelFormat::read(ctx, &field(ctx, dst, "format")?)?;
    if !sf.byte_channels() || sf.bytes != df.bytes || sf.masks != df.masks
        || df.palette.is_some()
    {
        return None;
    }
    let dimension = |object, key| usize::try_from(field(ctx, object, key)?.as_i64()).ok();
    let width = dimension(src, "w")?;
    let height = dimension(src, "h")?;
    let row = width.checked_mul(sf.bytes)?;
    let pitch = dimension(src, "pitch")?;
    if dimension(dst, "w")? != width || dimension(dst, "h")? != height
        || pitch < row || dimension(dst, "pitch")? != pitch
        || (packed && pitch != row)
    {
        return None;
    }
    if row == 0 || height == 0 {
        return None;
    }
    let length = u32::try_from((height - 1).checked_mul(pitch)?.checked_add(row)?).ok()?;
    let row = u32::try_from(row).ok()?;
    let source = linear_memory32_address(&field(ctx, src, "pixels")?)?;
    let destination = linear_memory32_address(&field(ctx, dst, "pixels")?)?;
    ctx.with_linear_memory(|memory| {
        [source, destination].into_iter().all(|address| {
            u64::from(address) + u64::from(length) <= memory.len() as u64
        }).then_some((row, length))
    })?
}

fn linear_memory32_address(value: &Value) -> Option<u32> {
    match value {
        Value::I32(value) => u32::try_from(*value).ok(),
        Value::I64(value) => u32::try_from(*value).ok(),
        Value::F64(value) if value.is_finite() && *value >= 0.0
            && *value <= f64::from(u32::MAX) && value.trunc() == *value => Some(*value as u32),
        _ => None,
    }
}

#[cfg(test)]
mod bulk_blit_proof_tests {
    use super::*;

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut object = Object::new();
        object.properties.extend(fields.iter().map(|(name, value)| ((*name).into(), value.clone())));
        Value::Object(heap::alloc(object))
    }

    #[test]
    fn complete_linear_blit_matches_snapshot_conversion() {
        for bytes in [3, 4] {
            let format = object(&[
                ("BitsPerPixel", Value::I32(bytes * 8)),
                ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
                ("Bmask", Value::I64(16711680)),
                ("Amask", Value::I64(if bytes == 4 { 4278190080 } else { 0 })),
            ]);
            for offset in [-1, 0, 1, 64] {
                let source = object(&[("format", format.clone()), ("w", Value::I32(4)),
                    ("h", Value::I32(3)), ("pitch", Value::I32(4 * bytes)),
                    ("pixels", Value::I32(64))]);
                let dest = object(&[("format", format.clone()), ("w", Value::I32(4)),
                    ("h", Value::I32(3)), ("pitch", Value::I32(4 * bytes)),
                    ("pixels", Value::I32(64 + offset))]);
                let args = [source, Value::Null, dest, Value::Null, Value::I32(0)];
                let mut memory: Vec<u8> = (0..256).map(|i| (i * 37 + 129) as u8).collect();
                let mut expected = memory.clone();
                expected.copy_within(64..64 + (12 * bytes) as usize, (64 + offset) as usize);
                let mut ctx = HostContext::empty();
                ctx.memory = Some(&mut memory);
                assert_eq!(bulk_blit_length(&ctx, &args), Some((12 * bytes) as u32));
                assert_eq!(convert_blit(&mut ctx, &args), Some(()));
                assert_eq!(memory, expected);
            }
        }
    }

    #[test]
    fn raw_copy_requires_disjoint_complete_byte_channels() {
        for (bytes, masks, accepted) in [
            (4, [0xff, 0xff00, 0xff0000, 0xff000000], true),
            (4, [0xff0000, 0xff00, 0xff, 0xff000000], true),
            (3, [0xff, 0xff00, 0xff0000, 0], true),
            (4, [0xff, 0xff00, 0xff0000, 0], false),
            (4, [0xff, 0xff, 0xff0000, 0xff000000], false),
            (2, [0xf800, 0x7e0, 0x1f, 0], false),
            (4, [0x3ff00000, 0xffc00, 0x3ff, 0xc0000000], false),
        ] {
            let format = PixelFormat { bytes, masks, palette: None };
            assert_eq!(format.byte_channels(), accepted);
            if accepted {
                // Identity must include extrema and the unsigned high byte.
                for byte in 0..=255 {
                    let mut output = [0; 4];
                    format.encode([byte; 4], &mut output[..bytes]).unwrap();
                    assert_eq!(&output[..bytes], &vec![byte; bytes]);
                }
            }
        }
        assert!(!PixelFormat { bytes: 4,
            masks: [0xff, 0xff00, 0xff0000, 0xff000000],
            palette: Some(Value::Null) }.byte_channels());
    }

    #[test]
    fn bulk_addresses_are_exact_memory32_words_without_saturation() {
        for value in [Value::I32(-1), Value::I64(-1), Value::F64(-1.0),
            Value::F64(1.5), Value::F64(f64::NAN), Value::F64(f64::INFINITY),
            Value::F64(4294967296.0), Value::Null] {
            assert_eq!(linear_memory32_address(&value), None);
        }
        assert_eq!(linear_memory32_address(&Value::F64(4294967295.0)), Some(u32::MAX));
        assert_eq!(linear_memory32_address(&Value::I64(4294967295)), Some(u32::MAX));
        assert_eq!(linear_memory32_address(&Value::I32(0)), Some(0));
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
    let source = field(ctx, src, "pixels")?;
    let destination = field(ctx, dst, "pixels")?;
    if sf.bytes == 1 && df.bytes == 4 && df.byte_channels()
        && let Some((0, colors)) = &palette
        && colors.len() >= 256
    {
        // Full four-byte coverage proves every component has one destination
        // byte. Reorder the small palette once, rather than encode each pixel.
        let mut ordered = [[0u8; 4]; 256];
        let colors = if df.masks == [0xff, 0xff00, 0xff0000, 0xff000000] {
            colors.as_slice()
        } else {
            for (color, pixel) in colors.iter().take(256).zip(&mut ordered) {
                for component in 0..4 {
                    pixel[(df.masks[component].trailing_zeros() / 8) as usize] = color[component];
                }
            }
            &ordered[..]
        };
        let source_offset = sy.checked_mul(sp)?.checked_add(sx)?;
        let destination_offset = dy.checked_mul(dp)?.checked_add(dx.checked_mul(4)?)?;
        let converted = if tw == w && th == h {
            expand_linear_palette_rows(ctx, &source, &destination,
                source_offset, destination_offset, w, h, sp, dp, 0, colors)
        } else {
            scale_linear_palette_rows(ctx, &source, &destination,
                source_offset, destination_offset, w, h, sp, tw, th, dp, colors)
        };
        if converted.is_some() { return Some(()); }
    }
    if tw == w && th == h && sf.bytes == 1
        && direct_rgba_format(&df, tw, tw.checked_mul(4)?)
        && let Some((first, colors)) = &palette
        && expand_linear_palette_rows(ctx, &source, &destination,
            sy.checked_mul(sp)?.checked_add(sx)?,
            dy.checked_mul(dp)?.checked_add(dx.checked_mul(4)?)?,
            w, h, sp, dp, *first, colors).is_some()
    {
        return Some(());
    }
    if tw == w && th == h
        && sf.palette.is_none() && df.palette.is_none()
        && sf.byte_channels() && df.byte_channels()
        && sf.bytes == df.bytes && sf.masks == df.masks
        && copy_linear_matching_rows(ctx, &source, &destination,
            sy.checked_mul(sp)?.checked_add(sx.checked_mul(sf.bytes)?)?,
            dy.checked_mul(dp)?.checked_add(dx.checked_mul(df.bytes)?)?,
            w.checked_mul(sf.bytes)?, h, sp, dp).is_some()
    {
        return Some(());
    }
    if sf.bytes == 4
        && df.bytes == 4
        && sf.byte_channels()
        && df.byte_channels()
        && sf.masks == df.masks
        && scale_linear_rgba_rows(ctx, &source, &destination,
            sy.checked_mul(sp)?.checked_add(sx.checked_mul(4)?)?,
            dy.checked_mul(dp)?.checked_add(dx.checked_mul(4)?)?,
            w, h, sp, tw, th, dp).is_some()
    {
        return Some(());
    }
    if tw == w && th == h
        && sf.palette.is_none() && df.palette.is_none()
        && sf.byte_channels() && df.byte_channels()
        && (sf.bytes != df.bytes || sf.masks != df.masks)
        && convert_linear_byte_rows(ctx, &source, &destination,
            sy.checked_mul(sp)?.checked_add(sx.checked_mul(sf.bytes)?)?,
            dy.checked_mul(dp)?.checked_add(dx.checked_mul(df.bytes)?)?,
            w, h, sp, dp, &sf, &df).is_some()
    {
        return Some(());
    }
    let input = read_span(ctx, Some(&source), input_len)?;
    let mut output = read_span(ctx, Some(&destination), output_len)?;
    if sf.bytes == 1 && df.bytes == 4 && df.byte_channels()
        && let Some((0, colors)) = &palette
        && colors.len() >= 256
    {
        // Managed surfaces and overlapping linear views arrive here as
        // independent snapshots. Reuse the same palette scaler as the direct
        // linear path, with its table reordered once into destination bytes.
        let mut ordered = [[0u8; 4]; 256];
        for (color, pixel) in colors.iter().take(256).zip(&mut ordered) {
            for component in 0..4 {
                pixel[(df.masks[component].trailing_zeros() / 8) as usize] = color[component];
            }
        }
        let source_offset = sy.checked_mul(sp)?.checked_add(sx)?;
        let destination_offset = dy.checked_mul(dp)?.checked_add(dx.checked_mul(4)?)?;
        if scale_palette_bytes(
            input.get(source_offset..)?, output.get_mut(destination_offset..)?,
            w, h, sp, tw, th, dp, &ordered,
        ).is_some() {
            return write_span(ctx, &destination, &output);
        }
    }
    if sf.palette.is_none() && df.palette.is_none()
        && sf.byte_channels() && df.byte_channels()
        && sf.bytes == df.bytes && sf.masks == df.masks
    {
        // The snapshots retain source-at-entry semantics for managed storage,
        // C views and overlapping linear regions. Complete byte masks prove
        // matching formats round-trip identically without pixel decoding.
        let source_offset = sy.checked_mul(sp)?.checked_add(sx.checked_mul(sf.bytes)?)?;
        let destination_offset = dy.checked_mul(dp)?.checked_add(dx.checked_mul(df.bytes)?)?;
        let input_rows = input.get(source_offset..)?;
        let output_rows = output.get_mut(destination_offset..)?;
        let copied = if tw == w && th == h {
            frame_kernels::copy_snapshot_byte_rows(
                input_rows, output_rows, w.checked_mul(sf.bytes)?, h, sp, dp,
            )
        } else if sf.bytes == 4 {
            scale_rgba_bytes(input_rows, output_rows, w, h, sp, tw, th, dp)
        } else {
            None
        };
        if copied.is_some() {
            return write_span(ctx, &destination, &output);
        }
    }
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

fn scale_rgba_bytes(
    input: &[u8], output: &mut [u8],
    width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0 { return None; }
    let source_row_bytes = width.checked_mul(4)?;
    let output_row_bytes = output_width.checked_mul(4)?;
    if pitch < source_row_bytes || output_pitch < output_row_bytes { return None; }
    let source_len = (rows - 1).checked_mul(pitch)?.checked_add(source_row_bytes)?;
    let output_len = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_row_bytes)?;
    if input.len() < source_len || output.len() < output_len { return None; }
    if width.checked_mul(3) == Some(output_width) {
        return frame_kernels::scale_integer_rgba_bytes::<3>(
            input, output, width, rows, pitch, output_rows, output_pitch,
        );
    }
    if width.checked_mul(4) == Some(output_width) {
        return frame_kernels::scale_integer_rgba_bytes::<4>(
            input, output, width, rows, pitch, output_rows, output_pitch,
        );
    }
    if width.checked_mul(2) == Some(output_width) {
        return frame_kernels::scale_double_rgba_bytes(
            input, output, width, rows, pitch, output_rows, output_pitch,
        );
    }
    // Select once, outside the row loop. Keep the original validated loop
    // in this function for downscaling and equal-width copies.
    if output_rows > rows && width != output_width {
        return scale_rgba_enlarged_rows(input, output, width, rows, pitch,
            output_width, output_rows, output_pitch);
    }
    let (x_step, x_remainder) = (width / output_width, width % output_width);
    let (y_step, y_remainder) = (rows / output_rows, rows % output_rows);
    let (x_threshold, y_threshold) = (output_width - x_remainder, output_rows - y_remainder);
    let (mut source_y, mut y_fraction) = (0, 0);
    for y in 0..output_rows {
        let row = &input[source_y * pitch..source_y * pitch + source_row_bytes];
        let pixels = &mut output[y * output_pitch..y * output_pitch + output_row_bytes];
        if width == output_width {
            pixels.copy_from_slice(row);
        } else {
            let (mut source_x, mut x_fraction) = (0, 0);
            for pixel in pixels.chunks_exact_mut(4) {
                pixel.copy_from_slice(&row[source_x * 4..source_x * 4 + 4]);
                source_x += x_step;
                if x_fraction >= x_threshold {
                    x_fraction -= x_threshold;
                    source_x += 1;
                } else { x_fraction += x_remainder; }
            }
        }
        source_y += y_step;
        if y_fraction >= y_threshold {
            y_fraction -= y_threshold;
            source_y += 1;
        } else { y_fraction += y_remainder; }
    }
    Some(())
}

// Called only after scale_rgba_bytes validates dimensions and both spans.
fn scale_rgba_enlarged_rows(
    input: &[u8], output: &mut [u8],
    width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
) -> Option<()> {
    let source_row_bytes = width * 4;
    let output_row_bytes = output_width * 4;
    let (x_step, x_remainder) = (width / output_width, width % output_width);
    let (y_step, y_remainder) = (rows / output_rows, rows % output_rows);
    let (x_threshold, y_threshold) = (output_width - x_remainder, output_rows - y_remainder);
    let (mut source_y, mut y_fraction) = (0, 0);
    let mut previous_source_y = None;
    for y in 0..output_rows {
        if previous_source_y == Some(source_y) {
            // Nearest-neighbor enlargement repeats complete rows. Reuse only
            // pixel bytes, leaving destination padding untouched. All spans
            // were checked above; no guest pointer or unchecked access is used.
            let previous = (y - 1) * output_pitch;
            output.copy_within(previous..previous + output_row_bytes, y * output_pitch);
        } else {
        let row = &input[source_y * pitch..source_y * pitch + source_row_bytes];
        let pixels = &mut output[y * output_pitch..y * output_pitch + output_row_bytes];
        if width == output_width {
            pixels.copy_from_slice(row);
        } else {
            let (mut source_x, mut x_fraction) = (0, 0);
            for pixel in pixels.chunks_exact_mut(4) {
                pixel.copy_from_slice(&row[source_x * 4..source_x * 4 + 4]);
                source_x += x_step;
                // Incremental floor(x * width / output_width). Compare before
                // adding the remainder to avoid coordinate multiplication,
                // per-pixel division, and fractional-accumulator overflow.
                if x_fraction >= x_threshold {
                    x_fraction -= x_threshold;
                    source_x += 1;
                } else { x_fraction += x_remainder; }
            }
        }
        }
        previous_source_y = Some(source_y);
        source_y += y_step;
        if y_fraction >= y_threshold {
            y_fraction -= y_threshold;
            source_y += 1;
        } else { y_fraction += y_remainder; }
    }
    Some(())
}

fn scale_linear_rgba_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0 { return None; }
    let source_row_bytes = width.checked_mul(4)?;
    let output_row_bytes = output_width.checked_mul(4)?;
    if pitch < source_row_bytes || output_pitch < output_row_bytes { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize).checked_add(destination_offset)?;
    let source_len = (rows - 1).checked_mul(pitch)?.checked_add(source_row_bytes)?;
    let output_len = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_row_bytes)?;
    let source_end = source.checked_add(source_len)?;
    let destination_end = destination.checked_add(output_len)?;
    // Preserve snapshot semantics for every possibly overlapping layout.
    if source < destination_end && destination < source_end { return None; }
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source_end <= destination {
            let (input, output) = memory.split_at_mut(destination);
            scale_rgba_bytes(&input[source..source_end], &mut output[..output_len],
                width, rows, pitch, output_width, output_rows, output_pitch)
        } else {
            let (output, input) = memory.split_at_mut(source);
            scale_rgba_bytes(&input[..source_len], &mut output[destination..destination_end],
                width, rows, pitch, output_width, output_rows, output_pitch)
        }
    })?
}

fn copy_linear_matching_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    row_bytes: usize, rows: usize, source_pitch: usize, destination_pitch: usize,
) -> Option<()> {
    if source_pitch == destination_pitch {
        return copy_linear_rgba_rows(ctx, source, destination, source_offset,
            destination_offset, row_bytes, rows, source_pitch);
    }
    if rows == 0 || row_bytes == 0 { return Some(()); }
    if source_pitch < row_bytes || destination_pitch < row_bytes { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize)
        .checked_add(destination_offset)?;
    let source_span = (rows - 1).checked_mul(source_pitch)?.checked_add(row_bytes)?;
    let destination_span = (rows - 1).checked_mul(destination_pitch)?.checked_add(row_bytes)?;
    let source_end = source.checked_add(source_span)?;
    let destination_end = destination.checked_add(destination_span)?;
    // With unequal pitches, choosing a single row direction is insufficient.
    // Reject possibly overlapping spans before any write; retain snapshots.
    if source < destination_end && destination < source_end { return None; }
    let copy = |input: &[u8], output: &mut [u8]| {
        for (src, dst) in input.chunks(source_pitch).zip(output.chunks_mut(destination_pitch)) {
            dst[..row_bytes].copy_from_slice(&src[..row_bytes]);
        }
    };
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source_end <= destination {
            let (input, output) = memory.split_at_mut(destination);
            copy(&input[source..source_end], &mut output[..destination_span]);
        } else {
            let (output, input) = memory.split_at_mut(source);
            copy(&input[..source_span], &mut output[destination..destination_end]);
        }
        Some(())
    })?
}

// Equal pitches make every row's displacement identical. Copy in address
// order (reversed when moving upward in memory), preserving snapshot semantics
// even when one destination row overlaps a different source row. Row padding
// is never copied. No guest pointer escapes the checked memory borrow.
fn copy_linear_rgba_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    row_bytes: usize, rows: usize, pitch: usize,
) -> Option<()> {
    if rows == 0 || row_bytes == 0 { return Some(()); }
    if row_bytes > pitch { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize)
        .checked_add(destination_offset)?;
    let span = (rows - 1).checked_mul(pitch)?.checked_add(row_bytes)?;
    let source_end = source.checked_add(span)?;
    let destination_end = destination.checked_add(span)?;
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source == destination { return Some(()); }
        if pitch == row_bytes && (source_end <= destination || destination_end <= source) {
            memory.copy_within(source..source_end, destination);
            return Some(());
        }
        for step in 0..rows {
            let row = if destination > source { rows - 1 - step } else { step };
            let offset = row * pitch;
            memory.copy_within(source + offset..source + offset + row_bytes,
                destination + offset);
        }
        Some(())
    })?
}

fn convert_linear_byte_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    width: usize, rows: usize, source_pitch: usize, destination_pitch: usize,
    source_format: &PixelFormat, destination_format: &PixelFormat,
) -> Option<()> {
    if !source_format.byte_channels() || !destination_format.byte_channels()
        || source_format.palette.is_some() || destination_format.palette.is_some()
        || width == 0 || rows == 0 { return None; }
    let source_bytes = width.checked_mul(source_format.bytes)?;
    let destination_bytes = width.checked_mul(destination_format.bytes)?;
    if source_pitch < source_bytes || destination_pitch < destination_bytes { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize)
        .checked_add(destination_offset)?;
    let source_span = (rows - 1).checked_mul(source_pitch)?.checked_add(source_bytes)?;
    let destination_span = (rows - 1).checked_mul(destination_pitch)?.checked_add(destination_bytes)?;
    let source_end = source.checked_add(source_span)?;
    let destination_end = destination.checked_add(destination_span)?;
    if source < destination_end && destination < source_end { return None; }
    // Match the generic decoder/encoder: missing RGB and unused output bytes
    // are zero, while missing alpha is opaque only if the output stores alpha.
    let (mapping, defaults) = frame_kernels::byte_channel_mapping(
        source_format.masks, destination_format.masks,
    );
    let convert = |input: &[u8], output: &mut [u8]| {
        for (src, dst) in input.chunks(source_pitch).zip(output.chunks_mut(destination_pitch)) {
            for (src_pixel, dst_pixel) in src[..source_bytes].chunks_exact(source_format.bytes)
                .zip(dst[..destination_bytes].chunks_exact_mut(destination_format.bytes)) {
                for (byte, value) in dst_pixel.iter_mut().enumerate() {
                    *value = mapping[byte].map_or(defaults[byte], |index| src_pixel[index]);
                }
            }
        }
    };
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source_end <= destination {
            let (input, output) = memory.split_at_mut(destination);
            convert(&input[source..source_end], &mut output[..destination_span]);
        } else {
            let (output, input) = memory.split_at_mut(source);
            convert(&input[..source_span], &mut output[destination..destination_end]);
        }
        Some(())
    })?
}

fn scale_palette_bytes(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
    colors: &[[u8; 4]],
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0
        || colors.len() < 256 { return None; }
    let output_bytes = output_width.checked_mul(4)?;
    if pitch < width || output_pitch < output_bytes { return None; }
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(width)?;
    let output_span = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    if input.len() < source_span || output.len() < output_span { return None; }
    let double_width = width.checked_mul(2) == Some(output_width);
    let (xs, xr) = (width / output_width, width % output_width);
    let (ys, yr) = (rows / output_rows, rows % output_rows);
    let (xt, yt) = (output_width - xr, output_rows - yr);
    let (mut source_y, mut fraction_y) = (0, 0);
    let mut previous_y = None;
    for y in 0..output_rows {
        if previous_y == Some(source_y) {
            let previous = (y - 1) * output_pitch;
            output.copy_within(previous..previous + output_bytes, y * output_pitch);
        } else {
            let row = &input[source_y * pitch..source_y * pitch + width];
            let output_row = &mut output[y * output_pitch..y * output_pitch + output_bytes];
            if double_width {
                frame_kernels::double_palette_row(row, output_row, colors)?;
            } else {
                let (mut source_x, mut fraction_x) = (0, 0);
                for pixel in output_row.chunks_exact_mut(4) {
                    pixel.copy_from_slice(&colors[usize::from(row[source_x])]);
                    source_x += xs;
                    if fraction_x >= xt {
                        fraction_x -= xt; source_x += 1;
                    } else { fraction_x += xr; }
                }
            }
        }
        previous_y = Some(source_y);
        source_y += ys;
        if fraction_y >= yt {
            fraction_y -= yt; source_y += 1;
        } else { fraction_y += yr; }
    }
    Some(())
}

fn scale_linear_palette_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    width: usize, rows: usize, pitch: usize,
    output_width: usize, output_rows: usize, output_pitch: usize,
    colors: &[[u8; 4]],
) -> Option<()> {
    if width == 0 || rows == 0 || output_width == 0 || output_rows == 0
        || colors.len() < 256 { return None; }
    let output_bytes = output_width.checked_mul(4)?;
    if pitch < width || output_pitch < output_bytes { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize)
        .checked_add(destination_offset)?;
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(width)?;
    let output_span = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    let source_end = source.checked_add(source_span)?;
    let destination_end = destination.checked_add(output_span)?;
    if source < destination_end && destination < source_end { return None; }
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source_end <= destination {
            let (input, output) = memory.split_at_mut(destination);
            scale_palette_bytes(&input[source..source_end], &mut output[..output_span],
                width, rows, pitch, output_width, output_rows, output_pitch, colors)
        } else {
            let (output, input) = memory.split_at_mut(source);
            scale_palette_bytes(&input[..source_span], &mut output[destination..destination_end],
                width, rows, pitch, output_width, output_rows, output_pitch, colors)
        }
    })?
}

fn expand_linear_palette_rows(
    ctx: &mut HostContext<'_>, source: &Value, destination: &Value,
    source_offset: usize, destination_offset: usize,
    width: usize, rows: usize, source_pitch: usize, destination_pitch: usize,
    first: usize, colors: &[[u8; 4]],
) -> Option<()> {
    if rows == 0 || width == 0 { return Some(()); }
    let row_bytes = width.checked_mul(4)?;
    if source_pitch < width || destination_pitch < row_bytes { return None; }
    let source = (linear_memory32_address(source)? as usize).checked_add(source_offset)?;
    let destination = (linear_memory32_address(destination)? as usize)
        .checked_add(destination_offset)?;
    let source_span = (rows - 1).checked_mul(source_pitch)?.checked_add(width)?;
    let destination_span = (rows - 1).checked_mul(destination_pitch)?.checked_add(row_bytes)?;
    let source_end = source.checked_add(source_span)?;
    let destination_end = destination.checked_add(destination_span)?;
    // Expansion can overwrite yet-unread indices when buffers overlap.
    // Preserve the existing snapshot fallback for all such layouts.
    if source < destination_end && destination < source_end { return None; }
    let expand = |input: &[u8], output: &mut [u8]| -> Option<()> {
        // The snapshot converter commits only after every palette lookup
        // succeeds. Retain that behavior before writing directly to memory.
        // A zero-based 256-color palette covers the entire u8 domain;
        // lookup failure is impossible and no prevalidation scan is needed.
        if first != 0 || colors.len() < 256 {
            for row in 0..rows {
                for &index in &input[row * source_pitch..row * source_pitch + width] {
                    colors.get(usize::from(index).checked_sub(first)?)?;
                }
            }
        }
        for row in 0..rows {
            let indices = &input[row * source_pitch..row * source_pitch + width];
            let pixels = &mut output[row * destination_pitch..row * destination_pitch + row_bytes];
            for (&index, pixel) in indices.iter().zip(pixels.chunks_exact_mut(4)) {
                pixel.copy_from_slice(&colors[usize::from(index) - first]);
            }
        }
        Some(())
    };
    ctx.with_linear_memory_mut(|memory| {
        if source_end > memory.len() || destination_end > memory.len() { return None; }
        if source_end <= destination {
            let (input, output) = memory.split_at_mut(destination);
            expand(&input[source..source_end], &mut output[..destination_span])
        } else {
            let (output, input) = memory.split_at_mut(source);
            expand(&input[..source_span], &mut output[destination..destination_end])
        }
    })?
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
    let bytes = ta.buffer.lock().unwrap();
    typed_array_value_from_bytes(ta, &bytes, index)
}

fn typed_array_value_from_bytes(ta: &TypedArrayState, bytes: &[u8], index: usize) -> Option<Value> {
    let bpe = ta.elem.bytes_per_element();
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
        ObjectKind::TypedArray(ta) => {
            // Primitive numeric decoding cannot call back into the VM. Borrow
            // the backing storage once, retaining the same checked element
            // reads and zero fallback for incomplete/out-of-range elements.
            let bytes = ta.buffer.lock().unwrap();
            Some((0..ta.length)
                .map(|i| {
                    typed_array_value_from_bytes(ta, &bytes, i)
                        .map(|v| v.as_i32().clamp(0, 255) as u8)
                        .unwrap_or(0)
                })
                .collect())
        }
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

#[cfg(test)]
mod typed_snapshot_tests {
    use super::*;

    fn legacy_snapshot(value: &Value) -> Vec<u8> {
        let Value::Object(object) = value else { unreachable!() };
        let object = object.lock().unwrap();
        let ObjectKind::TypedArray(array) = &object.kind else { unreachable!() };
        if matches!(array.elem, TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped) {
            let bytes = array.buffer.lock().unwrap();
            let start = array.byte_offset.min(bytes.len());
            let end = start.saturating_add(array.length).min(bytes.len());
            return bytes[start..end].to_vec();
        }
        (0..array.length).map(|index| typed_array_value(array, index)
            .map(|value| value.as_i32().clamp(0, 255) as u8).unwrap_or(0)).collect()
    }

    fn source(bytes: Vec<u8>, elem: TypedElemKind, offset: usize, count: usize) -> Value {
        let value = typed_byte_array(bytes, elem, "snapshot-test");
        let Value::Object(object) = &value else { unreachable!() };
        {
            let mut object = object.lock().unwrap();
            let ObjectKind::TypedArray(array) = &mut object.kind else { unreachable!() };
            array.byte_offset = offset;
            array.length = count;
        }
        value
    }

    #[test]
    fn borrowed_snapshot_preserves_all_numeric_formats_and_truncated_views() {
        for elem in [TypedElemKind::I8, TypedElemKind::U8, TypedElemKind::U8Clamped,
            TypedElemKind::I16, TypedElemKind::U16, TypedElemKind::I32, TypedElemKind::U32,
            TypedElemKind::F32, TypedElemKind::F64, TypedElemKind::BigI64, TypedElemKind::BigU64] {
            for offset in [0, 1, 3, 127, usize::MAX] {
                for count in [0, 1, 2, 31] {
                    let bytes: Vec<_> = (0..127).map(|index| (index * 73 + 91) as u8).collect();
                    let value = source(bytes.clone(), elem, offset, count);
                    assert_eq!(bytes_from_value(None, Some(&value)), Some(legacy_snapshot(&value)));
                    let Value::Object(object) = &value else { unreachable!() };
                    let object = object.lock().unwrap();
                    let ObjectKind::TypedArray(array) = &object.kind else { unreachable!() };
                    assert_eq!(*array.buffer.lock().unwrap(), bytes);
                }
            }
        }
    }

    #[test]
    #[ignore = "paired non-byte SDL snapshot locking benchmark"]
    fn borrowed_snapshot_benchmark() {
        use std::{hint::black_box, time::Instant};
        for count in [320usize, 320 * 200] {
            let value = source(vec![1; count * 2], TypedElemKind::U16, 0, count);
            let measure = |bulk| {
                let start = Instant::now();
                for _ in 0..16 {
                    let output = if bulk { bytes_from_value(None, Some(black_box(&value))).unwrap() }
                        else { legacy_snapshot(black_box(&value)) };
                    black_box(output);
                }
                start.elapsed().as_secs_f64() * 1e6 / 16.0
            };
            let mut old = Vec::new(); let mut new = Vec::new();
            for trial in 0..21 {
                if trial % 2 == 0 { old.push(measure(false)); new.push(measure(true)); }
                else { new.push(measure(true)); old.push(measure(false)); }
            }
            old.sort_by(f64::total_cmp); new.sort_by(f64::total_cmp);
            println!("typed-snapshot count={count} per_element_us={:.3} bulk_us={:.3} speedup={:.3}",
                old[10], new[10], old[10] / new[10]);
        }
    }
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
