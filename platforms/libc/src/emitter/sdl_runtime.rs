use std::sync::{Arc, Mutex};

use vybe_runtime::heap;
use vybe_runtime::value::{ArrayBufferState, Object, ObjectKind, TypedArrayState, TypedElemKind};
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "libc:sdl",
        "rgbaImageData",
        Box::new(|ctx, args| {
            let width = int_arg(args, 1).max(0) as usize;
            let height = int_arg(args, 2).max(0) as usize;
            let len = width.saturating_mul(height).saturating_mul(4);
            let mut bytes = bytes_from_value(Some(ctx), args.first()).unwrap_or_default();
            bytes.resize(len, 0);
            bytes.truncate(len);
            image_data(bytes, width, height)
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
    object_field(ctx, &o, "__sdl_colors").or_else(|| palette.cloned())
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
        elem: TypedElemKind::U8Clamped,
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
    ta.properties.insert(
        "__type".into(),
        Value::String(Arc::from("Uint8ClampedArray")),
    );
    ta.properties.insert(
        "tostringtag".into(),
        Value::String(Arc::from("Uint8ClampedArray")),
    );
    Value::Object(heap::alloc(ta))
}
