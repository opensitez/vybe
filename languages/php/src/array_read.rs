//! PHP indexed reads. Kept in one host operation so an array lookup does not
//! expand into a full ArrayAccess/undefined AST at every source occurrence.
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

fn is_array_object(object: &Object) -> bool {
    matches!(
        object.properties.get("__php_array_object"),
        Some(Value::Bool(true))
    )
}

fn numeric_sort_value(value: &Value) -> Option<f64> {
    match value {
        Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) => {
            let number = value.as_f64();
            number.is_finite().then_some(number)
        }
        _ => None,
    }
}

fn sort_numeric_values_in_place(value: &Value, descending: bool) -> bool {
    let Value::Object(object) = value else {
        return false;
    };
    let mut object = object.lock().unwrap();
    match &mut object.kind {
        ObjectKind::Array(values) => {
            if !values
                .iter()
                .all(|value| numeric_sort_value(value).is_some())
            {
                return false;
            }
            values.sort_by(|left, right| {
                let order = numeric_sort_value(left)
                    .unwrap()
                    .total_cmp(&numeric_sort_value(right).unwrap());
                if descending { order.reverse() } else { order }
            });
            true
        }
        ObjectKind::Map(entries) => {
            if !entries
                .values()
                .all(|value| numeric_sort_value(value).is_some())
            {
                return false;
            }
            let mut sorted: Vec<_> = std::mem::take(entries).into_iter().collect();
            sorted.sort_by(|left, right| {
                let order = numeric_sort_value(&left.1)
                    .unwrap()
                    .total_cmp(&numeric_sort_value(&right.1).unwrap());
                if descending { order.reverse() } else { order }
            });
            entries.extend(sorted);
            true
        }
        _ => false,
    }
}

fn export_scalar(value: &Value) -> String {
    match value {
        Value::Null | Value::Undefined => "NULL".to_string(),
        Value::Bool(value) => if *value { "true" } else { "false" }.to_string(),
        Value::String(value) => format!("'{}'", value.replace('\\', "\\\\").replace('\'', "\\'")),
        Value::I32(value) => value.to_string(),
        Value::I64(value) => value.to_string(),
        Value::F64(value) => value.to_string(),
        _ => "NULL".to_string(),
    }
}

fn export_array_value(value: &Value, depth: usize, active: &mut Vec<usize>) -> String {
    let Value::Object(object) = value else {
        return export_scalar(value);
    };
    let identity = Arc::as_ptr(object) as usize;
    if active.contains(&identity) {
        return "NULL".to_string();
    }
    let entries = {
        let object = object.lock().unwrap();
        if is_array_object(&object) {
            return "NULL".to_string();
        }
        match &object.kind {
            ObjectKind::Map(entries) => entries
                .iter()
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect::<Vec<_>>(),
            ObjectKind::Array(items) => {
                let holes = match object.properties.get("__holes") {
                    Some(Value::Object(holes)) => match &holes.lock().unwrap().kind {
                        ObjectKind::Array(holes) => holes.clone(),
                        _ => Vec::new(),
                    },
                    _ => Vec::new(),
                };
                items
                    .iter()
                    .enumerate()
                    .filter(|(index, _)| {
                        !holes.iter().any(
                            |hole| matches!(hole, Value::I32(value) if *value == *index as i32),
                        )
                    })
                    .map(|(index, value)| (Value::I64(index as i64), value.clone()))
                    .collect::<Vec<_>>()
            }
            _ => return "NULL".to_string(),
        }
    };
    active.push(identity);
    let mut output = String::from("array (\n");
    for (key, item) in entries {
        output.push_str(&"  ".repeat(depth + 1));
        output.push_str(&export_scalar(&key));
        output.push_str(" => ");
        output.push_str(&export_array_value(&item, depth + 1, active));
        output.push_str(",\n");
    }
    output.push_str(&"  ".repeat(depth));
    output.push(')');
    active.pop();
    output
}

fn array_entry_count(object: &Object) -> Option<usize> {
    if is_array_object(object) {
        return None;
    }
    match &object.kind {
        ObjectKind::Map(entries) => Some(entries.len()),
        ObjectKind::Array(items) => {
            let Some(Value::Object(holes)) = object.properties.get("__holes") else {
                return Some(items.len());
            };
            let holes = holes.lock().unwrap();
            let ObjectKind::Array(holes) = &holes.kind else {
                return Some(items.len());
            };
            Some(
                (0..items.len())
                    .filter(|index| {
                        !holes.iter().any(
                            |hole| matches!(hole, Value::I32(value) if *value == *index as i32),
                        )
                    })
                    .count(),
            )
        }
        _ => None,
    }
}

// Copy PHP array values in one runtime operation. Objects (including reference
// cells) keep their identity. Only the active recursion path is memoized: two
// separate occurrences of a nested value must become independent array values.
fn copy_array(value: &Value, active: &mut Vec<(usize, Value)>) -> Value {
    let Value::Object(source) = value else {
        return value.clone();
    };
    let identity = Arc::as_ptr(source) as usize;
    if let Some((_, copy)) = active.iter().find(|(id, _)| *id == identity) {
        return copy.clone();
    }
    let mut snapshot = {
        let source = source.lock().unwrap();
        if !matches!(source.kind, ObjectKind::Array(_) | ObjectKind::Map(_))
            || is_array_object(&source)
        {
            return value.clone();
        }
        source.clone()
    };
    let target = heap::alloc(Object::new());
    let result = Value::Object(target.clone());
    active.push((identity, result.clone()));
    match &mut snapshot.kind {
        ObjectKind::Array(items) => {
            for item in items {
                *item = copy_array(item, active);
            }
        }
        ObjectKind::Map(items) => {
            for item in items.values_mut() {
                *item = copy_array(item, active);
            }
        }
        _ => unreachable!(),
    }
    active.pop();
    // Preserve array metadata, including its internal cursor. Metadata is not
    // user array contents and must not recursively clone prototypes/callables.
    *target.lock().unwrap() = snapshot;
    result
}

fn array_map_snapshot(value: &Value) -> Option<Object> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    if is_array_object(&object) {
        return None;
    }
    match &object.kind {
        ObjectKind::Map(_) => Some(object.clone()),
        ObjectKind::Array(items) => {
            let mut result = Object::new();
            result.kind = ObjectKind::Map(Default::default());
            let ObjectKind::Map(entries) = &mut result.kind else {
                unreachable!()
            };
            for (index, value) in items.iter().enumerate() {
                entries.insert(Value::I64(index as i64), value.clone());
            }
            Some(result)
        }
        _ => None,
    }
}

// Replacement descends only when both values are PHP arrays. Objects and
// reference cells preserve their identity; ordinary nested array values copy.
fn replace_arrays_recursive(base: &Value, overlay: &Value) -> Option<Value> {
    let base = array_map_snapshot(base)?;
    let overlay = array_map_snapshot(overlay)?;
    let (ObjectKind::Map(base), ObjectKind::Map(overlay)) = (base.kind, overlay.kind) else {
        unreachable!()
    };
    let mut merged = base.clone();
    for (key, value) in &overlay {
        merged.insert(key.clone(), value.clone());
    }
    for (key, value) in &mut merged {
        *value = match (base.get(key), overlay.get(key)) {
            (Some(base), Some(overlay)) => replace_arrays_recursive(base, overlay)
                .unwrap_or_else(|| copy_array(overlay, &mut Vec::new())),
            (_, Some(overlay)) => copy_array(overlay, &mut Vec::new()),
            (Some(base), None) => copy_array(base, &mut Vec::new()),
            _ => unreachable!(),
        };
    }
    let packed = merged.keys().enumerate().all(|(index, key)| {
        matches!(key, Value::I32(_) | Value::I64(_) | Value::F64(_)) && key.as_f64() == index as f64
    });
    let mut result = Object::new();
    result.kind = if packed {
        ObjectKind::Array(merged.into_values().collect())
    } else {
        ObjectKind::Map(merged)
    };
    Some(Value::Object(heap::alloc(result)))
}

// JSON.parse produces ordinary objects. PHP's associative decode requires
// maps recursively, including objects nested in JSON arrays. The input is a
// freshly parsed tree, so converting its storage in place cannot affect aliases.
fn json_to_arrays(value: &Value) {
    let Value::Object(object) = value else { return };
    let children = {
        let mut object = object.lock().unwrap();
        if matches!(object.kind, ObjectKind::Ordinary) {
            let keys = match object.properties.get("__keys") {
                Some(Value::Object(keys)) => match &keys.lock().unwrap().kind {
                    ObjectKind::Array(keys) => keys
                        .iter()
                        .filter_map(|key| {
                            if let Value::String(key) = key {
                                Some(key.to_string())
                            } else {
                                None
                            }
                        })
                        .collect::<Vec<_>>(),
                    _ => Vec::new(),
                },
                _ => object.properties.keys().cloned().collect(),
            };
            let mut map = ObjectKind::Map(Default::default());
            if let ObjectKind::Map(entries) = &mut map {
                for key in keys {
                    if let Some(value) = object.properties.shift_remove(&key) {
                        let numeric = key.parse::<i64>().ok().filter(|n| n.to_string() == key);
                        let key = numeric
                            .map(Value::I64)
                            .unwrap_or_else(|| Value::String(Arc::from(key)));
                        entries.insert(key, value);
                    }
                }
            }
            object.properties.clear();
            object.kind = map;
        }
        match &object.kind {
            ObjectKind::Array(items) => items.clone(),
            ObjectKind::Map(items) => items.values().cloned().collect(),
            _ => Vec::new(),
        }
    };
    for child in children {
        json_to_arrays(&child);
    }
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:property",
        "isInaccessible",
        Box::new(|_, args| {
            let [
                Value::Object(instance),
                Value::String(field),
                Value::String(scope),
                ..,
            ] = args
            else {
                return Value::Bool(false);
            };
            let constructor = instance.lock().unwrap().get("constructor");
            let Value::Object(constructor) = constructor else {
                return Value::Bool(false);
            };
            let constructor = constructor.lock().unwrap();
            let owner = constructor.get("__php_lsb_class").to_string();
            let Value::Object(visibility) = constructor.get("__php_magic_visibility") else {
                return Value::Bool(false);
            };
            let visibility = visibility.lock().unwrap();
            let ObjectKind::Map(fields) = &visibility.kind else {
                return Value::Bool(false);
            };
            let Some(level) = fields.get(&Value::String(field.clone())) else {
                return Value::Bool(false);
            };
            Value::Bool(level.to_string() != "public" && scope.as_ref() != owner)
        }),
    );
    fw.register_host_fn(
        "php:array",
        "varExport",
        Box::new(|_, args| {
            Value::String(Arc::from(export_array_value(
                args.first().unwrap_or(&Value::Null),
                0,
                &mut Vec::new(),
            )))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "unpackValues",
        Box::new(|_, args| {
            let Some(Value::Object(object)) = args.first() else {
                return Value::Null;
            };
            let values = match &object.lock().unwrap().kind {
                ObjectKind::Array(values) => values.clone(),
                ObjectKind::Map(entries) => entries.values().cloned().collect(),
                _ => return Value::Null,
            };
            let mut result = Object::new();
            result.kind = ObjectKind::Array(values);
            Value::Object(heap::alloc(result))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "countEntries",
        Box::new(|_, args| {
            let Some(Value::Object(object)) = args.first() else {
                return Value::Null;
            };
            array_entry_count(&object.lock().unwrap())
                .map(|count| Value::I64(count as i64))
                .unwrap_or(Value::Null)
        }),
    );
    fw.register_host_fn(
        "php:value",
        "isEmpty",
        Box::new(|_, args| {
            let empty = match args.first() {
                None | Some(Value::Null | Value::Undefined) => true,
                Some(Value::Bool(value)) => !value,
                Some(Value::I32(value)) => *value == 0,
                Some(Value::I64(value)) => *value == 0,
                Some(Value::F64(value)) => *value == 0.0,
                Some(Value::String(value)) => value.is_empty() || value.as_ref() == "0",
                Some(Value::Object(object)) => {
                    let object = object.lock().unwrap();
                    if is_array_object(&object) {
                        false
                    } else {
                        match &object.kind {
                            ObjectKind::Map(entries) => entries.is_empty(),
                            ObjectKind::Array(items) => {
                                if items.is_empty() {
                                    true
                                } else if let Some(Value::Object(holes)) =
                                    object.properties.get("__holes")
                                {
                                    let holes = holes.lock().unwrap();
                                    if let ObjectKind::Array(holes) = &holes.kind {
                                        (0..items.len()).all(|index| {
                                            holes.iter().any(|hole| {
                                        matches!(hole, Value::I32(value) if *value == index as i32)
                                    })
                                        })
                                    } else {
                                        false
                                    }
                                } else {
                                    false
                                }
                            }
                            _ => false,
                        }
                    }
                }
                _ => false,
            };
            Value::Bool(empty)
        }),
    );
    fw.register_host_fn(
        "php:array",
        "replaceRecursive",
        Box::new(|_, args| {
            let [base, overlay, ..] = args else {
                return Value::Null;
            };
            replace_arrays_recursive(base, overlay).unwrap_or(Value::Null)
        }),
    );
    fw.register_host_fn("php:array", "isArray", Box::new(|_, args| {
        let is_array = if let Some(Value::Object(object)) = args.first() {
            let object = object.lock().unwrap();
            matches!(object.kind, ObjectKind::Array(_) | ObjectKind::Map(_))
                && !is_array_object(&object)
                && !matches!(object.properties.get("__kind"), Some(Value::String(kind)) if kind.as_ref() == "object")
        } else { false };
        // Internal predicate consumed by typed bytecode conditions. The public
        // PHP is_array adapter converts this to a PHP boolean value.
        Value::I32(i32::from(is_array))
    }));
    fw.register_host_fn(
        "php:array",
        "containsStrictString",
        Box::new(|_, args| {
            let [
                Value::Object(array),
                Value::String(needle),
                Value::Bool(true),
                ..,
            ] = args
            else {
                return Value::Null;
            };
            let object = array.lock().unwrap();
            if is_array_object(&object) {
                return Value::Null;
            }
            let matches = |value: &Value| matches!(value, Value::String(text) if text == needle);
            match &object.kind {
                ObjectKind::Array(items) => Value::Bool(items.iter().any(matches)),
                ObjectKind::Map(items) => Value::Bool(items.values().any(matches)),
                _ => Value::Null,
            }
        }),
    );
    fw.register_host_fn(
        "php:array",
        "isArrayObject",
        Box::new(|_, args| {
            Value::Bool(matches!(args.first(), Some(Value::Object(object)) if {
                let object = object.lock().unwrap();
                is_array_object(&object) || object.properties.contains_key("__php_array_storage")
            }))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "newArrayObject",
        Box::new(|_, args| {
            let input = args.first().cloned().unwrap_or(Value::Null);
            let output = copy_array(&input, &mut Vec::new());
            let mut object = match output {
                Value::Object(object) => object.lock().unwrap().clone(),
                _ => Object::new_array(Vec::new()),
            };
            object
                .properties
                .insert("__php_array_object".into(), Value::Bool(true));
            object
                .properties
                .insert("__type".into(), Value::String(Arc::from("ArrayObject")));
            Value::Object(heap::alloc(object))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "objectExchange",
        Box::new(|_, args| {
            let Some(Value::Object(receiver)) = args.first() else {
                return Value::Null;
            };
            let replacement = args.get(1).cloned().unwrap_or(Value::Null);
            let replacement = copy_array(&replacement, &mut Vec::new());
            let replacement_kind = match &replacement {
                Value::Object(value) => value.lock().unwrap().kind.clone(),
                _ => ObjectKind::Array(Vec::new()),
            };
            let old = {
                let mut object = receiver.lock().unwrap();
                if is_array_object(&object) {
                    let mut snapshot = object.clone();
                    snapshot.properties.shift_remove("__php_array_object");
                    snapshot.properties.shift_remove("__type");
                    object.kind = replacement_kind;
                    Value::Object(heap::alloc(snapshot))
                } else {
                    object
                        .properties
                        .insert("__php_array_storage".into(), replacement)
                        .unwrap_or_else(|| {
                            Value::Object(heap::alloc(Object::new_array(Vec::new())))
                        })
                }
            };
            copy_array(&old, &mut Vec::new())
        }),
    );
    fw.register_host_fn(
        "php:array",
        "objectCopy",
        Box::new(|_, args| {
            let Some(Value::Object(receiver)) = args.first() else {
                return Value::Null;
            };
            let storage = {
                let object = receiver.lock().unwrap();
                if is_array_object(&object) {
                    let mut snapshot = object.clone();
                    snapshot.properties.shift_remove("__php_array_object");
                    snapshot.properties.shift_remove("__type");
                    Value::Object(heap::alloc(snapshot))
                } else {
                    object
                        .properties
                        .get("__php_array_storage")
                        .cloned()
                        .unwrap_or_else(|| {
                            Value::Object(heap::alloc(Object::new_array(Vec::new())))
                        })
                }
            };
            copy_array(&storage, &mut Vec::new())
        }),
    );
    fw.register_host_fn(
        "php:array",
        "objectCount",
        Box::new(|_, args| {
            let Some(Value::Object(receiver)) = args.first() else {
                return Value::Null;
            };
            let storage = receiver
                .lock()
                .unwrap()
                .properties
                .get("__php_array_storage")
                .cloned();
            let Some(Value::Object(storage)) = storage else {
                return Value::Null;
            };
            array_entry_count(&storage.lock().unwrap())
                .map(|count| Value::F64(count as f64))
                .unwrap_or(Value::Null)
        }),
    );
    fw.register_host_fn(
        "php:array",
        "compactForSort",
        Box::new(|_, args| {
            let value = args.first().cloned().unwrap_or(Value::Null);
            if let Value::Object(object) = &value {
                let mut object = object.lock().unwrap();
                if !is_array_object(&object) {
                    let entries = match &object.kind {
                        ObjectKind::Array(items) => Some(
                            items
                                .iter()
                                .filter(|item| !matches!(item, Value::Undefined))
                                .cloned()
                                .collect(),
                        ),
                        ObjectKind::Map(items) => Some(
                            items
                                .values()
                                .filter(|item| !matches!(item, Value::Undefined))
                                .cloned()
                                .collect(),
                        ),
                        _ => None,
                    };
                    if let Some(entries) = entries {
                        object.kind = ObjectKind::Array(entries);
                        object.properties.shift_remove("__holes");
                    }
                }
            }
            value
        }),
    );
    fw.register_host_fn(
        "php:array",
        "newClosure",
        Box::new(|_, args| {
            let Some(Value::Object(function)) = args.first() else {
                return Value::Null;
            };
            // Wasm ref.func is interned. A PHP closure expression creates a new
            // object even without captures; captured reference cells stay shared.
            let closure = function.lock().unwrap().clone();
            Value::Object(heap::alloc(closure))
        }),
    );

    // Object and closure Values retain their Arc across assignment and calls.
    // The allocation identity stays stable while alive; PHP allows IDs to be
    // reused after an object is destroyed.
    for operation in ["objectId", "objectHash"] {
        fw.register_host_fn(
            "php:array",
            operation,
            Box::new(move |_, args| {
                let Some(Value::Object(object)) = args.first() else {
                    return Value::Null;
                };
                let identity = Arc::as_ptr(object) as usize as u64;
                if operation == "objectId" {
                    Value::bigint_i64(identity as i64)
                } else {
                    Value::String(Arc::from(format!("{identity:032x}")))
                }
            }),
        );
    }

    fw.register_host_fn(
        "php:array",
        "jsonDepth",
        Box::new(|_, args| {
            let text = args.first().map(ToString::to_string).unwrap_or_default();
            let (mut quoted, mut escaped, mut depth, mut maximum) = (false, false, 0i32, 0i32);
            for byte in text.bytes() {
                if quoted {
                    if escaped {
                        escaped = false;
                    } else if byte == b'\\' {
                        escaped = true;
                    } else if byte == b'"' {
                        quoted = false;
                    }
                } else {
                    match byte {
                        b'"' => quoted = true,
                        b'[' | b'{' => {
                            depth += 1;
                            maximum = maximum.max(depth);
                        }
                        b']' | b'}' => depth -= 1,
                        _ => {}
                    }
                }
            }
            Value::I32(maximum + 1)
        }),
    );
    fw.register_host_fn(
        "php:array",
        "fromJson",
        Box::new(|_, args| {
            let result = args.first().cloned().unwrap_or(Value::Null);
            json_to_arrays(&result);
            result
        }),
    );
    fw.register_host_fn(
        "php:array",
        "copy",
        Box::new(|_, args| {
            args.first()
                .map_or(Value::Null, |value| copy_array(value, &mut Vec::new()))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "copyCursor",
        Box::new(|_, args| {
            let [Value::Object(source), Value::Object(target), ..] = args else {
                return Value::Undefined;
            };
            let cursor = source
                .lock()
                .unwrap()
                .properties
                .get("__php_array_cursor")
                .cloned();
            if let Some(cursor) = cursor {
                target
                    .lock()
                    .unwrap()
                    .properties
                    .insert("__php_array_cursor".into(), cursor);
            }
            Value::Undefined
        }),
    );
    // PHP's internal array cursor belongs to the array value, independently
    // of foreach iterators. Keep it outside the elements so keys/values and
    // serialization do not expose it as an array entry.
    for operation in ["current", "next", "reset", "end", "prev", "key"] {
        fw.register_host_fn(
            "php:array",
            operation,
            Box::new(move |_, args| {
                let missing = if operation == "key" {
                    Value::Null
                } else {
                    Value::Bool(false)
                };
                let Some(Value::Object(array)) = args.first() else {
                    return missing;
                };
                let mut object = array.lock().unwrap();
                let len = match &object.kind {
                    ObjectKind::Array(items) => items.len(),
                    ObjectKind::Map(items) => items.len(),
                    _ => return missing,
                } as i64;
                let cursor = object
                    .properties
                    .get("__php_array_cursor")
                    .map_or(0, |value| value.as_f64() as i64);
                let cursor = match operation {
                    "reset" => 0,
                    "end" => len - 1,
                    "next" if cursor >= 0 && cursor < len => cursor + 1,
                    "prev" if cursor >= 0 && cursor < len => cursor - 1,
                    _ => cursor,
                };
                if !matches!(operation, "current" | "key") {
                    object
                        .properties
                        .insert("__php_array_cursor".into(), Value::I64(cursor));
                }
                if cursor < 0 || cursor >= len {
                    return missing;
                }
                match &object.kind {
                    ObjectKind::Array(items) => {
                        if operation == "key" {
                            Value::I64(cursor)
                        } else {
                            items[cursor as usize].clone()
                        }
                    }
                    ObjectKind::Map(items) => {
                        let (key, value) = items.get_index(cursor as usize).unwrap();
                        if operation == "key" {
                            key.clone()
                        } else {
                            value.clone()
                        }
                    }
                    _ => missing,
                }
            }),
        );
    }
    fw.register_host_fn(
        "php:array",
        "sortNumericValues",
        Box::new(|_, args| {
            let value = args.first().unwrap_or(&Value::Null);
            let descending = matches!(args.get(1), Some(Value::Bool(true)));
            Value::Bool(sort_numeric_values_in_place(value, descending))
        }),
    );
    fw.register_host_fn(
        "php:array",
        "addDeclaredFields",
        Box::new(|ctx, args| {
            let [source, Value::Object(result), ..] = args else {
                return Value::Null;
            };
            let Value::Object(source) = source else {
                return Value::Object(result.clone());
            };
            let fields = {
                let source = source.lock().unwrap();
                ctx.declared_fields(source.type_id)
                    .into_iter()
                    .enumerate()
                    .filter(|(_, (name, _))| !name.starts_with("__"))
                    .filter_map(|(index, (name, _))| {
                        source.fields.get(index).cloned().map(|value| (name, value))
                    })
                    .collect::<Vec<_>>()
            };
            let mut result_object = result.lock().unwrap();
            if let ObjectKind::Map(entries) = &mut result_object.kind {
                for (name, value) in fields {
                    entries
                        .entry(Value::String(Arc::from(name)))
                        .or_insert(value);
                }
            }
            Value::Object(result.clone())
        }),
    );
    fw.register_host_fn(
        "php:array",
        "local_set",
        Box::new(|_, args| {
            let [Value::Object(array), key, value, ..] = args else {
                return Value::Null;
            };
            if matches!(key, Value::Null | Value::Undefined) {
                return Value::Null;
            }
            let mut object = array.lock().unwrap();
            match &mut object.kind {
                ObjectKind::Map(entries) => {
                    entries.insert(key.clone(), value.clone());
                }
                ObjectKind::Array(items) => {
                    // A dense integer write keeps the indexed representation.
                    let dense_index = match key {
                        Value::I32(n) if *n >= 0 => *n as usize,
                        Value::I64(n) if *n >= 0 => *n as usize,
                        Value::F64(n)
                            if n.is_finite()
                                && *n >= 0.0
                                && n.fract() == 0.0
                                && *n <= items.len() as f64 =>
                        {
                            *n as usize
                        }
                        _ => usize::MAX,
                    };
                    if dense_index < items.len() {
                        items[dense_index] = value.clone();
                    } else if dense_index == items.len() {
                        items.push(value.clone());
                    } else {
                        // A PHP array can gain a string or sparse numeric key.
                        // Convert its existing indexed entries in place so this
                        // local write does not depend on a separate assignment.
                        let old = std::mem::take(items);
                        let mut map = ObjectKind::Map(Default::default());
                        if let ObjectKind::Map(entries) = &mut map {
                            for (index, item) in old.into_iter().enumerate() {
                                entries.insert(Value::I64(index as i64), item);
                            }
                            entries.insert(key.clone(), value.clone());
                        }
                        object.kind = map;
                    }
                }
                // PHP dynamic property writes also arrive as indexed places. A
                // property bag must keep its identity and other fields; falling
                // through to array_replace turns it into a one-entry array.
                ObjectKind::Ordinary => {
                    object.set(key.to_string(), value.clone());
                }
                _ => return Value::Null,
            }
            Value::Object(array.clone())
        }),
    );
    fw.register_host_fn(
        "php:array",
        "local_append",
        Box::new(|_, args| {
            let [Value::Object(array), value, ..] = args else {
                return Value::Null;
            };
            let mut object = array.lock().unwrap();
            match &mut object.kind {
                ObjectKind::Array(items) => items.push(value.clone()),
                ObjectKind::Map(entries) => {
                    let next = entries
                        .keys()
                        .filter_map(|key| match key {
                            Value::I32(n) if *n >= 0 => Some(*n as i64),
                            Value::I64(n) if *n >= 0 => Some(*n),
                            Value::F64(n)
                                if n.is_finite()
                                    && *n >= 0.0
                                    && n.fract() == 0.0
                                    && *n < i64::MAX as f64 =>
                            {
                                Some(*n as i64)
                            }
                            _ => None,
                        })
                        .max()
                        .unwrap_or(-1)
                        .saturating_add(1);
                    entries.insert(Value::I64(next), value.clone());
                }
                _ => return Value::Null,
            }
            Value::Object(array.clone())
        }),
    );
    fw.register_host_fn(
        "php:array",
        "nested_member_set",
        Box::new(|ctx, args| {
            let [
                Value::Object(owner),
                Value::String(field),
                bucket_key,
                leaf_key,
                value,
                ..,
            ] = args
            else {
                return Value::Null;
            };
            if matches!(bucket_key, Value::Null | Value::Undefined) {
                return Value::Null;
            }
            let outer = {
                let object = owner.lock().unwrap();
                object.properties.get(field.as_ref()).cloned().or_else(|| {
                    ctx.declared_fields(object.type_id)
                        .iter()
                        .position(|(name, _)| name == field.as_ref())
                        .and_then(|index| object.fields.get(index).cloned())
                })
            };
            let Some(Value::Object(outer)) = outer else {
                return Value::Null;
            };
            let bucket = {
                let mut outer_obj = outer.lock().unwrap();
                if let ObjectKind::Array(items) = &mut outer_obj.kind {
                    let indexed = std::mem::take(items);
                    outer_obj.kind = ObjectKind::Map(Default::default());
                    if let ObjectKind::Map(entries) = &mut outer_obj.kind {
                        for (index, item) in indexed.into_iter().enumerate() {
                            entries.insert(Value::I64(index as i64), item);
                        }
                    }
                }
                let ObjectKind::Map(entries) = &mut outer_obj.kind else {
                    return Value::Null;
                };
                match entries.get(bucket_key) {
                    Some(Value::Object(existing)) => existing.clone(),
                    Some(Value::Null | Value::Undefined) | None => {
                        let mut object = Object::new();
                        object.kind = ObjectKind::Map(Default::default());
                        let created = heap::alloc(object);
                        entries.insert(bucket_key.clone(), Value::Object(created.clone()));
                        created
                    }
                    _ => return Value::Null,
                }
            };
            {
                let mut bucket_obj = bucket.lock().unwrap();
                match &mut bucket_obj.kind {
                    ObjectKind::Map(entries) => {
                        let key = if matches!(leaf_key, Value::Null | Value::Undefined) {
                            let next = entries
                                .keys()
                                .filter_map(|key| match key {
                                    Value::I32(index) => Some(*index as i64),
                                    Value::I64(index) => Some(*index),
                                    _ => None,
                                })
                                .max()
                                .map_or(0, |index| index.saturating_add(1).max(0));
                            Value::I64(next)
                        } else {
                            leaf_key.clone()
                        };
                        entries.insert(key, value.clone());
                    }
                    ObjectKind::Array(items)
                        if matches!(leaf_key, Value::Null | Value::Undefined) =>
                    {
                        items.push(value.clone());
                    }
                    _ => return Value::Null,
                }
            }
            Value::Object(bucket)
        }),
    );
    fw.register_host_fn(
        "php:array",
        "read",
        Box::new(|ctx, args| {
            let object = args.first().cloned().unwrap_or(Value::Null);
            let key = args.get(1).cloned().unwrap_or(Value::Null);
            match &object {
                Value::Object(cell) => {
                    let method = {
                        let object = cell.lock().unwrap();
                        let method = object.get("offsetGet");
                        match &method {
                            Value::Object(f)
                                if matches!(
                                    f.lock().unwrap().kind,
                                    ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                                ) =>
                            {
                                Some(method)
                            }
                            _ => None,
                        }
                    };
                    if let Some(method) = method {
                        return ctx.invoke_with_receiver(&method, object, &[key]);
                    }
                    let data = cell.lock().unwrap();
                    match &data.kind {
                        ObjectKind::Map(map) => {
                            let direct = map.get(&key).cloned().or_else(|| match &key {
                                Value::String(s) => s
                                    .parse::<i32>()
                                    .ok()
                                    .and_then(|n| map.get(&Value::I32(n)))
                                    .cloned(),
                                Value::I32(n) => {
                                    map.get(&Value::String(Arc::from(n.to_string()))).cloned()
                                }
                                _ => None,
                            });
                            direct.unwrap_or(Value::Null)
                        }
                        _ => data.get(&format!("{key}")),
                    }
                }
                Value::String(s) => {
                    let len = s.len() as i64;
                    let mut index = key.as_f64() as i64;
                    if index < 0 {
                        index += len;
                    }
                    usize::try_from(index)
                        .ok()
                        .and_then(|i| s.as_bytes().get(i))
                        .map(|byte| Value::String(Arc::from(char::from(*byte).to_string())))
                        .unwrap_or(Value::Null)
                }
                _ => Value::Null,
            }
        }),
    );
}
