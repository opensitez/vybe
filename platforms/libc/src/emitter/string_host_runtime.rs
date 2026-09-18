use vybe_runtime::value::{Object, ObjectKind, TypedArrayState, TypedElemKind};
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "libc:string",
        "charToStr",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Null);
            Value::String(c_string_from_value(Some(ctx), &value).into())
        }),
    );
    vm.register_host_fn(
        "libc:string",
        "charPtrAdd",
        Box::new(|ctx, args| {
            let target = args.first().cloned().unwrap_or(Value::Null);
            let offset = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            char_ptr_add(Some(ctx), target, offset)
        }),
    );
    vm.register_host_fn(
        "libc:string",
        "charPtrWrite",
        Box::new(|ctx, args| {
            let Some(target) = args.first().cloned() else {
                return Value::Null;
            };
            let index = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let code = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let ch = args
                .get(3)
                .and_then(char_value)
                .or_else(|| char::from_u32(code.max(0) as u32))
                .unwrap_or('\0');
            char_ptr_write(Some(ctx), target, index, code, ch)
        }),
    );
    vm.register_host_fn(
        "libc:string",
        "strncpyCarray",
        Box::new(|ctx, args| {
            let Some(dest) = args.first().cloned() else {
                return Value::Null;
            };
            let src = args.get(1).cloned().unwrap_or(Value::Null);
            let n = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            strncpy_carray(Some(ctx), &dest, &src, n);
            dest
        }),
    );
    vm.register_host_fn(
        "libc:string",
        "isNullPtr",
        Box::new(|_ctx, args| {
            let is_null = matches!(
                args.first(),
                None | Some(Value::Null | Value::TypedNull(_) | Value::Undefined)
            );
            Value::Bool(is_null)
        }),
    );
    vm.register_host_fn(
        "libc:string",
        "strlenCString",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Null);
            Value::I32(c_string_len(Some(ctx), &value) as i32)
        }),
    );
}

fn char_value(value: &Value) -> Option<char> {
    match value {
        Value::String(s) => s.chars().next(),
        other => char::from_u32(other.as_i32().max(0) as u32),
    }
}

fn char_ptr_add(ctx: Option<&HostContext<'_>>, target: Value, offset: usize) -> Value {
    if let Some((base, base_offset)) = carray_view(ctx, &target) {
        return carray_ref(base, base_offset.saturating_add(offset));
    }
    if let Value::String(s) = target {
        let sliced: String = s.chars().skip(offset).collect();
        return Value::String(sliced.into());
    }
    target
}

fn char_ptr_write(
    ctx: Option<&HostContext<'_>>,
    target: Value,
    index: usize,
    code: i32,
    ch: char,
) -> Value {
    if let Some((base, offset)) = carray_view(ctx, &target) {
        write_indexed_value(&base, offset.saturating_add(index), Value::I32(code));
        return target;
    }
    if let Value::String(s) = target {
        let mut chars: Vec<char> = s.chars().collect();
        if chars.len() <= index {
            chars.resize(index + 1, '\0');
        }
        chars[index] = ch;
        let updated: String = chars.into_iter().collect();
        return Value::String(updated.into());
    }
    target
}

fn carray_ref(base: Value, idx: usize) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__ref_kind".into(), Value::String("carray".into()));
    obj.properties.insert("__base".into(), base);
    obj.properties.insert("__idx".into(), Value::I32(idx as i32));
    Value::Object(vybe_runtime::heap::alloc(obj))
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

fn strncpy_carray(ctx: Option<&HostContext<'_>>, dest: &Value, src: &Value, n: usize) {
    if n == 0 {
        return;
    }
    let Some((base, offset)) = carray_view(ctx, dest) else {
        return;
    };
    let text = c_string_from_value(ctx, src);
    let mut codes: Vec<Value> = text
        .chars()
        .take(n)
        .map(|ch| Value::I32(ch as i32))
        .collect();
    while codes.len() < n {
        codes.push(Value::I32(0));
    }

    let Value::Object(obj) = base else {
        return;
    };
    let mut o = obj.lock().unwrap();
    let mut array_len = None;
    match &mut o.kind {
        ObjectKind::Array(items) => {
            let required = offset.saturating_add(n);
            if items.len() < required {
                items.resize(required, Value::I32(0));
            }
            for (i, code) in codes.into_iter().enumerate() {
                items[offset + i] = code;
            }
            array_len = Some(items.len());
        }
        ObjectKind::TypedArray(ta) => write_typed_array(ta, offset, &codes),
        _ => {}
    }
    if let Some(len) = array_len {
        o.properties
            .insert("length".into(), Value::F64(len as f64));
    }
}

fn c_string_from_value(ctx: Option<&HostContext<'_>>, value: &Value) -> String {
    match value {
        Value::String(s) => s.split('\0').next().unwrap_or("").to_string(),
        Value::Object(_) => {
            if let Some((base, offset)) = carray_view(ctx, value) {
                return c_string_from_indexed(ctx, &base, offset);
            }
            c_string_from_indexed(ctx, value, 0)
        }
        Value::Null | Value::TypedNull(_) | Value::Undefined => String::new(),
        other => format!("{}", other),
    }
}

fn c_string_len(ctx: Option<&HostContext<'_>>, value: &Value) -> usize {
    match value {
        Value::String(s) => s.find('\0').unwrap_or(s.len()),
        Value::Object(_) => {
            if let Some((base, offset)) = carray_view(ctx, value) {
                return c_indexed_string_len(ctx, &base, offset);
            }
            c_indexed_string_len(ctx, value, 0)
        }
        _ => 0,
    }
}

fn c_indexed_string_len(ctx: Option<&HostContext<'_>>, value: &Value, offset: usize) -> usize {
    let len = indexed_len(value).unwrap_or(0);
    let mut count = 0;
    for i in offset..len {
        let Some(item) = indexed_value(ctx, value, i) else {
            break;
        };
        match item {
            Value::String(s) => {
                if s.is_empty() || s.as_ref() == "\0" {
                    break;
                }
            }
            other => {
                if other.as_i32() == 0 {
                    break;
                }
            }
        }
        count += 1;
    }
    count
}

fn c_string_from_indexed(ctx: Option<&HostContext<'_>>, value: &Value, offset: usize) -> String {
    let mut out = String::new();
    let len = indexed_len(value).unwrap_or(0);
    for i in offset..len {
        let Some(item) = indexed_value(ctx, value, i) else {
            break;
        };
        match item {
            Value::String(s) => {
                if s.is_empty() || s.as_ref() == "\0" {
                    break;
                }
                out.push_str(&s);
            }
            other => {
                let code = other.as_i32();
                if code == 0 {
                    break;
                }
                if let Some(ch) = char::from_u32(code as u32) {
                    out.push(ch);
                }
            }
        }
    }
    out
}

fn indexed_len(value: &Value) -> Option<usize> {
    let Value::Object(obj) = value else {
        return None;
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(items) => Some(items.len()),
        ObjectKind::TypedArray(ta) => Some(ta.length),
        _ => None,
    }
}

fn indexed_value(ctx: Option<&HostContext<'_>>, value: &Value, index: usize) -> Option<Value> {
    if let Some((base, offset)) = carray_view(ctx, value) {
        return indexed_value(ctx, &base, offset.saturating_add(index));
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

fn write_indexed_value(value: &Value, index: usize, item: Value) {
    let Value::Object(obj) = value else {
        return;
    };
    let mut o = obj.lock().unwrap();
    let mut array_len = None;
    match &mut o.kind {
        ObjectKind::Array(items) => {
            if items.len() <= index {
                items.resize(index + 1, Value::I32(0));
            }
            items[index] = item;
            array_len = Some(items.len());
        }
        ObjectKind::TypedArray(ta) => write_typed_array(ta, index, &[item]),
        _ => {
            o.properties.insert(index.to_string(), item);
        }
    }
    if let Some(len) = array_len {
        o.properties
            .insert("length".into(), Value::F64(len as f64));
    }
}

fn write_typed_array(ta: &mut TypedArrayState, offset: usize, codes: &[Value]) {
    if !matches!(
        ta.elem,
        TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped
    ) {
        return;
    }
    let mut bytes = ta.buffer.lock().unwrap();
    let start = ta.byte_offset.saturating_add(offset);
    let max = ta.byte_offset.saturating_add(ta.length).min(bytes.len());
    for (i, code) in codes.iter().enumerate() {
        let idx = start.saturating_add(i);
        if idx >= max {
            break;
        }
        bytes[idx] = code.as_i32().clamp(0, 255) as u8;
    }
}
