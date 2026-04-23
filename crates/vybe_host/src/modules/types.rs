//! Built-in .NET types: DateTime, StringBuilder, List, Dictionary.
//! Each constructor creates an object with methods as HostFunctions.

use chrono::Local;
use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    register_datetime(vm);
    register_stringbuilder(vm);
    register_list(vm);
    register_dictionary(vm);
    register_process(vm);
    register_queue_stack(vm);
    register_timespan(vm);
    register_guid(vm);
    register_primitives(vm);
}

// ============================================================
// DateTime
// ============================================================

fn register_datetime(vm: &mut VM) {
    // DateTime.Now → creates a DateTime object from current time
    vm.register_host_fn("vybe:types", "dateTimeNow", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        make_datetime_from_epoch_secs_with_offset(epoch_secs(), local_utc_offset_seconds())
    }));

    // DateTime.UtcNow → creates a DateTime object from current UTC time
    vm.register_host_fn("vybe:types", "dateTimeUtcNow", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        make_datetime_from_epoch_secs_with_offset(epoch_secs(), 0)
    }));

    // DateTime.Parse(str) → parse a date string
    vm.register_host_fn("vybe:types", "dateTimeParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // Simplified: just store the string, parse on demand
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("DateTime")));
        obj.properties.insert("__raw".into(), Value::String(Arc::from(s.as_str())));
        obj.properties.insert("__epoch".into(), Value::F64(0.0)); // TODO: parse
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // New DateTime(year, month, day) or New DateTime(year, month, day, hour, min, sec)
    vm.register_host_fn("vybe:types", "dateTimeNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let _this = args.first(); // ignore this from New
        let year = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(2000);
        let month = args.get(2).map(|v| v.as_f64() as u64).unwrap_or(1);
        let day = args.get(3).map(|v| v.as_f64() as u64).unwrap_or(1);
        let hour = args.get(4).map(|v| v.as_f64() as u64).unwrap_or(0);
        let min = args.get(5).map(|v| v.as_f64() as u64).unwrap_or(0);
        let sec = args.get(6).map(|v| v.as_f64() as u64).unwrap_or(0);
        make_datetime_from_parts(year, month, day, hour, min, sec, 0)
    }));

    // Instance methods called via vybe:runtime/callMethod or directly
    vm.register_host_fn("vybe:types", "dateTimeAddDays", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        dt_add(args, 86400.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddHours", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        dt_add(args, 3600.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddMinutes", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        dt_add(args, 60.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddSeconds", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        dt_add(args, 1.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddMonths", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let offset_seconds = datetime_offset_seconds(&o);
            let months = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(0);
            let (y, m, d, h, min, s) = decompose_with_offset(epoch, offset_seconds);
            let total_months = y * 12 + m as i64 + months;
            let ny = total_months / 12;
            let nm = ((total_months % 12) + 12) % 12;
            let nm = if nm == 0 { 12 } else { nm as u64 };
            return make_datetime_from_parts(ny, nm, d, h, min, s, offset_seconds);
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddYears", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let offset_seconds = datetime_offset_seconds(&o);
            let years = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(0);
            let (y, m, d, h, min, s) = decompose_with_offset(epoch, offset_seconds);
            return make_datetime_from_parts(y + years, m, d, h, min, s, offset_seconds);
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "dateTimeToString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let (y, m, d, h, min, s) = decompose_with_offset(epoch, datetime_offset_seconds(&o));
            let fmt = args.get(1).map(value_to_plain_string);
            let rendered = format_datetime(y, m, d, h, min, s, fmt.as_deref());
            return Value::String(Arc::from(rendered));
        }
        Value::String(Arc::from(""))
    }));
    vm.register_host_fn("vybe:types", "dateTimeToShortDate", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let (y, m, d, _, _, _) = decompose_with_offset(epoch, datetime_offset_seconds(&o));
            return Value::String(Arc::from(format!("{:02}/{:02}/{:04}", m, d, y).as_str()));
        }
        Value::String(Arc::from(""))
    }));
}

// ============================================================
// StringBuilder
// ============================================================

fn register_stringbuilder(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "stringBuilderNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let _this = args.first();
        let initial = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("StringBuilder")));
        obj.properties.insert("__buffer".into(), Value::String(Arc::from(initial.as_str())));
        obj.properties.insert("length".into(), Value::F64(initial.len() as f64));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "sbAppend", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let new_buf = format!("{}{}", current, text);
            let len = new_buf.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Arc::from(new_buf.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null) // return this for chaining
    }));

    vm.register_host_fn("vybe:types", "sbAppendLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let new_buf = format!("{}{}\n", current, text);
            let len = new_buf.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Arc::from(new_buf.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbToString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return o.properties.get("__buffer").cloned().unwrap_or(Value::String(Arc::from("")));
        }
        Value::String(Arc::from(""))
    }));

    vm.register_host_fn("vybe:types", "sbClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__buffer".into(), Value::String(Arc::from("")));
            o.properties.insert("length".into(), Value::F64(0.0));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbInsert", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let text = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            let mut current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let idx = index.min(current.len());
            current.insert_str(idx, &text);
            let len = current.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Arc::from(current.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbReplace", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let old = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let new = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let result = current.replace(&old, &new);
            let len = result.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Arc::from(result.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

// ============================================================
// List(Of T) — backed by array
// ============================================================

fn register_list(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "listNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let _this = args.first();
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // List.Add(item)
    vm.register_host_fn("vybe:types", "listAdd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.push(item);
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));

    // List.Remove(item) → bool
    vm.register_host_fn("vybe:types", "listRemove", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let before = elems.len();
                elems.retain(|e| format!("{}", e) != item_str);
                let removed = elems.len() < before;
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
                return Value::Bool(removed);
            }
        }
        Value::Bool(false)
    }));

    // List.RemoveAt(index)
    vm.register_host_fn("vybe:types", "listRemoveAt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if idx < elems.len() { elems.remove(idx); }
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));

    // List.Contains(item) → bool
    vm.register_host_fn("vybe:types", "listContains", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search));
            }
        }
        Value::Bool(false)
    }));

    // List.Count → number
    vm.register_host_fn("vybe:types", "listCount", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::F64(elems.len() as f64);
            }
        }
        Value::F64(0.0)
    }));

    // List.Clear()
    vm.register_host_fn("vybe:types", "listClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.clear();
                o.properties.insert("length".into(), Value::F64(0.0));
                o.properties.insert("count".into(), Value::F64(0.0));
            }
        }
        Value::Null
    }));

    // List.Item(index) → element at index
    vm.register_host_fn("vybe:types", "listItem", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.get(idx).cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // List.Insert(index, item)
    vm.register_host_fn("vybe:types", "listInsert", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let item = args.get(2).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let pos = idx.min(elems.len());
                elems.insert(pos, item);
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));

    // List.AddRange(collection)
    vm.register_host_fn("vybe:types", "listAddRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(1)) {
            let s = src.lock().unwrap();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items = src_elems.clone();
                drop(s);
                let mut d = dst.lock().unwrap();
                if let ObjectKind::Array(ref mut dst_elems) = d.kind {
                    dst_elems.extend(items);
                    let len = dst_elems.len() as f64;
                    d.properties.insert("length".into(), Value::F64(len));
                    d.properties.insert("count".into(), Value::F64(len));
                }
            }
        }
        Value::Null
    }));

    // List.Sort() — simple sort
    vm.register_host_fn("vybe:types", "listSort", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
            }
        }
        Value::Null
    }));

    // List.IndexOf(item[, start[, count]]) → index or -1
    vm.register_host_fn("vybe:types", "listIndexOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let start = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let end = if let Some(cnt) = args.get(3) {
                    (start + cnt.as_i32().max(0) as usize).min(elems.len())
                } else {
                    elems.len()
                };
                for (i, e) in elems[start.min(elems.len())..end].iter().enumerate() {
                    if format!("{}", e) == search { return Value::F64((start + i) as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));

    // List.LastIndexOf(item[, end]) → last index or -1. `end` is inclusive upper bound.
    vm.register_host_fn("vybe:types", "listLastIndexOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let end = args.get(2)
                    .map(|v| (v.as_i32() + 1).max(0) as usize)
                    .unwrap_or(elems.len());
                let end = end.min(elems.len());
                for (i, e) in elems[..end].iter().enumerate().rev() {
                    if format!("{}", e) == search { return Value::F64(i as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));

    // List.Reverse([index, count]) — full reverse or subrange
    vm.register_host_fn("vybe:types", "listReverse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if let Some(start_v) = args.get(1) {
                    let start = start_v.as_i32().max(0) as usize;
                    let count = args.get(2).map(|v| v.as_i32().max(0) as usize)
                        .unwrap_or_else(|| elems.len().saturating_sub(start));
                    let end = (start + count).min(elems.len());
                    elems[start..end].reverse();
                } else {
                    elems.reverse();
                }
            }
        }
        Value::Null
    }));

    // List.ToArray() / Clone() — return a shallow copy with __type = "List"
    vm.register_host_fn("vybe:types", "listClone", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let mut result = Object::new_array(elems.clone());
                result.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result)));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // List.ToArray() → copy (without __type for plain array return)
    vm.register_host_fn("vybe:types", "listToArray", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // List.InsertRange(index, collection)
    vm.register_host_fn("vybe:types", "listInsertRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(2)) {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let s = src.lock().unwrap();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items: Vec<Value> = src_elems.clone();
                drop(s);
                let mut d = dst.lock().unwrap();
                if let ObjectKind::Array(ref mut dst_elems) = d.kind {
                    let pos = idx.min(dst_elems.len());
                    for (i, item) in items.into_iter().enumerate() {
                        dst_elems.insert(pos + i, item);
                    }
                    let len = dst_elems.len() as f64;
                    d.properties.insert("length".into(), Value::F64(len));
                    d.properties.insert("count".into(), Value::F64(len));
                }
            }
        }
        Value::Null
    }));

    // List.RemoveRange(index, count)
    vm.register_host_fn("vybe:types", "listRemoveRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let count = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let start = idx.min(elems.len());
                let end = (start + count).min(elems.len());
                elems.drain(start..end);
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));

    // List.GetRange(index, count) → new ArrayList with sub-elements
    vm.register_host_fn("vybe:types", "listGetRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let count = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let start = idx.min(elems.len());
                let end = (start + count).min(elems.len());
                let sub: Vec<Value> = elems[start..end].to_vec();
                let mut result = Object::new_array(sub);
                result.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result)));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // List.SetRange(index, collection) — overwrite elements starting at index
    vm.register_host_fn("vybe:types", "listSetRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(2)) {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let s = src.lock().unwrap();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items: Vec<Value> = src_elems.clone();
                drop(s);
                let mut d = dst.lock().unwrap();
                if let ObjectKind::Array(ref mut dst_elems) = d.kind {
                    for (i, item) in items.into_iter().enumerate() {
                        let pos = idx + i;
                        if pos < dst_elems.len() {
                            dst_elems[pos] = item;
                        }
                    }
                }
            }
        }
        Value::Null
    }));

    // List.BinarySearch(value) — simplified to indexOf (assumes sorted, falls back to linear)
    vm.register_host_fn("vybe:types", "listBinarySearch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                for (i, e) in elems.iter().enumerate() {
                    if format!("{}", e) == search { return Value::F64(i as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));

    // List.Shift() — remove and return first element (for Queue TryDequeue)
    vm.register_host_fn("vybe:types", "listShift", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if !elems.is_empty() {
                    let first = elems.remove(0);
                    let len = elems.len() as f64;
                    o.properties.insert("length".into(), Value::F64(len));
                    o.properties.insert("count".into(), Value::F64(len));
                    return first;
                }
            }
        }
        Value::Null
    }));

    // List.Last() — return last element without removing (for Stack/Queue peek)
    vm.register_host_fn("vybe:types", "listLast", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.last().cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // List.Pop() — remove and return last element (for Stack TryPop)
    vm.register_host_fn("vybe:types", "listPop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if let Some(last) = elems.pop() {
                    let len = elems.len() as f64;
                    o.properties.insert("length".into(), Value::F64(len));
                    o.properties.insert("count".into(), Value::F64(len));
                    return last;
                }
            }
        }
        Value::Null
    }));
}

// ============================================================
// Dictionary(Of K, V)
// ============================================================

fn register_dictionary(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "dictNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Dictionary")));
        obj.properties.insert("__data".into(), Value::Object(Arc::new(Mutex::new(Object::new()))));
        obj.properties.insert("count".into(), Value::F64(0.0));
        obj.properties.insert("length".into(), Value::F64(0.0));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Dict.Add(key, value) / Dict.Item(key) = value
    vm.register_host_fn("vybe:types", "dictAdd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let data = {
                let o = obj.lock().unwrap();
                match o.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                let count = {
                    let mut data_obj = data.lock().unwrap();
                    data_obj.properties.insert(key.clone(), value.clone());
                    data_obj.properties.len() as f64
                };
                let mut outer = obj.lock().unwrap();
                if !key.starts_with("__") {
                    outer.properties.insert(key, value);
                }
                outer.properties.insert("count".into(), Value::F64(count));
                outer.properties.insert("length".into(), Value::F64(count));
            }
        }
        Value::Null
    }));

    // Dict.Item(key) → value
    vm.register_host_fn("vybe:types", "dictItem", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            // Direct property lookup (struct_new-based dicts)
            if let Some(val) = o.properties.get(&key) {
                return val.clone();
            }
        }
        Value::Null
    }));

    // Dict.TryGetValue(key[, outVal]) → value if found, null if not
    // The second arg (out-param) is ignored — out params aren't supported yet.
    // Returns the value directly so `found = dict.TryGetValue(key, out)` yields the value.
    vm.register_host_fn("vybe:types", "dictTryGetValue", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            if let Some(val) = o.properties.get(&key) {
                return val.clone();
            }
        }
        Value::Null
    }));

    // Dict.ContainsKey(key) → bool
    vm.register_host_fn("vybe:types", "dictContainsKey", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return Value::Bool(data.lock().unwrap().properties.contains_key(&key));
            }
            return Value::Bool(o.properties.contains_key(&key));
        }
        Value::Bool(false)
    }));

    // Dict.Remove(key) → bool
    vm.register_host_fn("vybe:types", "dictRemove", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let data = {
                let o = obj.lock().unwrap();
                match o.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                let (removed, count) = {
                    let mut data_obj = data.lock().unwrap();
                    let removed = data_obj.properties.remove(&key).is_some();
                    (removed, data_obj.properties.len() as f64)
                };
                let mut outer = obj.lock().unwrap();
                outer.properties.remove(&key);
                outer.properties.insert("count".into(), Value::F64(count));
                outer.properties.insert("length".into(), Value::F64(count));
                return Value::Bool(removed);
            }
        }
        Value::Bool(false)
    }));

    // Dict.Keys → array
    vm.register_host_fn("vybe:types", "dictKeys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            // Try __data first (old-style dicts), then enumerate properties directly
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let keys: Vec<Value> = data.lock().unwrap().properties.keys()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            // Direct property enumeration (struct_new-based dicts)
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| !k.starts_with("__")) // skip internal properties
                .map(|k| Value::String(Arc::from(k.as_str())))
                .collect();
            if !keys.is_empty() {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // Dict.Values → array
    vm.register_host_fn("vybe:types", "dictValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let vals: Vec<Value> = data.lock().unwrap().properties.values().cloned().collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
            let vals: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(_, v)| v.clone())
                .collect();
            if !vals.is_empty() {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // Dict.Clear()
    vm.register_host_fn("vybe:types", "dictClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let data = {
                let o = obj.lock().unwrap();
                match o.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                data.lock().unwrap().properties.clear();
                let mut outer = obj.lock().unwrap();
                outer.properties.retain(|k, _| k == "__type" || k == "__data" || k == "count" || k == "length");
                outer.properties.insert("count".into(), Value::F64(0.0));
                outer.properties.insert("length".into(), Value::F64(0.0));
            }
        }
        Value::Null
    }));
}

// ============================================================
// Process
// ============================================================

fn register_process(vm: &mut VM) {
    // ProcessStartInfo constructor: New ProcessStartInfo(cmd, args)
    vm.register_host_fn("vybe:types", "processStartInfoNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let cmd = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let cmd_args = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("ProcessStartInfo")));
        obj.properties.insert("filename".into(), Value::String(Arc::from(cmd.as_str())));
        obj.properties.insert("arguments".into(), Value::String(Arc::from(cmd_args.as_str())));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Process.Start(startInfo) — executes the process
    vm.register_host_fn("vybe:types", "processStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (cmd, cmd_args) = if let Some(Value::Object(si)) = args.first() {
            let o = si.lock().unwrap();
            let c = o.properties.get("filename").map(|v| format!("{}", v)).unwrap_or_default();
            let a = o.properties.get("arguments").map(|v| format!("{}", v)).unwrap_or_default();
            (c, a)
        } else {
            let c = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let a = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            (c, a)
        };
        match std::process::Command::new(&cmd).args(cmd_args.split_whitespace()).output() {
            Ok(output) => {
                let stdout = String::from_utf8_lossy(&output.stdout);
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Arc::from("Process")));
                obj.properties.insert("hasexited".into(), Value::Bool(true));
                obj.properties.insert("exitcode".into(), Value::F64(output.status.code().unwrap_or(-1) as f64));
                obj.properties.insert("standardoutput".into(), Value::String(Arc::from(stdout.as_ref())));
                Value::Object(Arc::new(Mutex::new(obj)))
            }
            Err(e) => {
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Arc::from("Process")));
                obj.properties.insert("hasexited".into(), Value::Bool(true));
                obj.properties.insert("exitcode".into(), Value::F64(-1.0));
                obj.properties.insert("error".into(), Value::String(Arc::from(format!("Error: {}", e).as_str())));
                Value::Object(Arc::new(Mutex::new(obj)))
            }
        }
    }));

    vm.register_host_fn("vybe:types", "processWaitForExit", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        // Process.Start waits for completion before returning the Process object.
        Value::Null
    }));

    // Process constructor (bare)
    vm.register_host_fn("vybe:types", "processNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Process")));
        obj.properties.insert("hasexited".into(), Value::Bool(false));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
}

// ============================================================
// DateTime helpers
// ============================================================

fn epoch_secs() -> u64 {
    super::clock::now_secs()
}

fn local_utc_offset_seconds() -> i32 {
    Local::now().offset().local_minus_utc()
}

fn make_datetime_from_epoch_secs_with_offset(secs: u64, offset_seconds: i32) -> Value {
    let (y, m, d, h, min, s) = decompose_with_offset(secs, offset_seconds);
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("DateTime")));
    obj.properties.insert("__epoch".into(), Value::F64(secs as f64));
    obj.properties.insert("__offset_seconds".into(), Value::F64(offset_seconds as f64));
    obj.properties.insert("year".into(), Value::F64(y as f64));
    obj.properties.insert("month".into(), Value::F64(m as f64));
    obj.properties.insert("day".into(), Value::F64(d as f64));
    obj.properties.insert("hour".into(), Value::F64(h as f64));
    obj.properties.insert("minute".into(), Value::F64(min as f64));
    obj.properties.insert("second".into(), Value::F64(s as f64));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_datetime_from_parts(year: i64, month: u64, day: u64, hour: u64, min: u64, sec: u64, offset_seconds: i32) -> Value {
    let epoch = date_to_epoch_with_offset(year, month, day, hour, min, sec, offset_seconds);
    make_datetime_from_epoch_secs_with_offset(epoch, offset_seconds)
}

fn datetime_offset_seconds(obj: &Object) -> i32 {
    obj.properties
        .get("__offset_seconds")
        .map(|v| v.as_f64() as i32)
        .unwrap_or(0)
}

fn dt_add(args: &[Value], multiplier: f64) -> Value {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.lock().unwrap();
        let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0);
        let offset_seconds = datetime_offset_seconds(&o);
        let amount = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let new_epoch = (epoch + amount * multiplier) as u64;
        return make_datetime_from_epoch_secs_with_offset(new_epoch, offset_seconds);
    }
    Value::Null
}

fn value_to_plain_string(value: &Value) -> String {
    match value {
        Value::String(s) => s.to_string(),
        _ => format!("{}", value),
    }
}

fn format_datetime(y: i64, m: u64, d: u64, h: u64, min: u64, s: u64, fmt: Option<&str>) -> String {
    match fmt.unwrap_or("") {
        "" => format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, m, d, h, min, s),
        "d" => format!("{:02}/{:02}/{:04}", m, d, y),
        "t" => format!("{:02}:{:02}", h, min),
        spec => {
            let hour12 = match h % 12 {
                0 => 12,
                v => v,
            };
            let am_pm = if h < 12 { "AM" } else { "PM" };
            let mut rendered = spec.to_string();
            for (token, replacement) in [
                ("yyyy", format!("{:04}", y)),
                ("MM", format!("{:02}", m)),
                ("dd", format!("{:02}", d)),
                ("HH", format!("{:02}", h)),
                ("hh", format!("{:02}", hour12)),
                ("mm", format!("{:02}", min)),
                ("ss", format!("{:02}", s)),
                ("tt", am_pm.to_string()),
            ] {
                rendered = rendered.replace(token, &replacement);
            }
            rendered
        }
    }
}

fn decompose(total_secs: u64) -> (i64, u64, u64, u64, u64, u64) {
    let days = total_secs / 86400;
    let tod = total_secs % 86400;
    let h = tod / 3600;
    let min = (tod % 3600) / 60;
    let s = tod % 60;
    let mut y = 1970i64;
    let mut rem = days as i64;
    loop {
        let diy = if y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) { 366 } else { 365 };
        if rem < diy { break; }
        rem -= diy;
        y += 1;
    }
    let leap = y % 4 == 0 && (y % 100 != 0 || y % 400 == 0);
    let md = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut m: u64 = 1;
    for dim in &md {
        if rem < *dim { break; }
        rem -= dim;
        m += 1;
    }
    (y, m, (rem + 1) as u64, h, min, s)
}

fn decompose_with_offset(total_secs: u64, offset_seconds: i32) -> (i64, u64, u64, u64, u64, u64) {
    let shifted = if offset_seconds >= 0 {
        total_secs.saturating_add(offset_seconds as u64)
    } else {
        total_secs.saturating_sub((-offset_seconds) as u64)
    };
    decompose(shifted)
}

fn date_to_epoch(year: i64, month: u64, day: u64, hour: u64, min: u64, sec: u64) -> u64 {
    let mut days: i64 = 0;
    for y in 1970..year {
        days += if y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) { 366 } else { 365 };
    }
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let md = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    for i in 0..(month.saturating_sub(1) as usize).min(12) {
        days += md[i];
    }
    days += day.saturating_sub(1) as i64;
    (days as u64) * 86400 + hour * 3600 + min * 60 + sec
}

fn date_to_epoch_with_offset(year: i64, month: u64, day: u64, hour: u64, min: u64, sec: u64, offset_seconds: i32) -> u64 {
    let base_epoch = date_to_epoch(year, month, day, hour, min, sec) as i128;
    (base_epoch - offset_seconds as i128).max(0) as u64
}

// ============================================================
// Queue / Stack / HashSet
// ============================================================

fn register_queue_stack(vm: &mut VM) {
    // Queue — backed by array, FIFO
    vm.register_host_fn("vybe:types", "queueNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("Queue")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "queueEnqueue", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.push(item);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "queueDequeue", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if !elems.is_empty() {
                    let val = elems.remove(0);
                    let len = elems.len() as f64;
                    o.properties.insert("count".into(), Value::F64(len));
                    return val;
                }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "queuePeek", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.first().cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "collectionPeek", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let is_stack = o.properties.get("__type")
                .and_then(|v| if let Value::String(s) = v { Some(s.as_ref() == "Stack") } else { None })
                .unwrap_or(false);
            if let ObjectKind::Array(ref elems) = o.kind {
                return if is_stack {
                    elems.last().cloned().unwrap_or(Value::Null)
                } else {
                    elems.first().cloned().unwrap_or(Value::Null)
                };
            }
        }
        Value::Null
    }));

    // Stack — backed by array, LIFO
    vm.register_host_fn("vybe:types", "stackNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("Stack")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "stackPush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.push(item);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "stackPop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let val = elems.pop().unwrap_or(Value::Null);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
                return val;
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "stackPeek", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.last().cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // HashSet — backed by array with uniqueness
    vm.register_host_fn("vybe:types", "hashSetNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("HashSet")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "hashSetAdd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let item_str = format!("{}", item);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                if !elems.iter().any(|e| format!("{}", e) == item_str) {
                    elems.push(item);
                    let len = elems.len() as f64;
                    o.properties.insert("count".into(), Value::F64(len));
                    return Value::Bool(true);
                }
            }
        }
        Value::Bool(false)
    }));
    vm.register_host_fn("vybe:types", "hashSetContains", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search));
            }
        }
        Value::Bool(false)
    }));
    vm.register_host_fn("vybe:types", "hashSetRemove", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let before = elems.len();
                elems.retain(|e| format!("{}", e) != search);
                let removed = elems.len() < before;
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
                return Value::Bool(removed);
            }
        }
        Value::Bool(false)
    }));
}

// ============================================================
// TimeSpan
// ============================================================

fn register_timespan(vm: &mut VM) {
    for (name, mult) in &[
        ("fromDays", 86400.0f64),
        ("fromHours", 3600.0),
        ("fromMinutes", 60.0),
        ("fromSeconds", 1.0),
        ("fromMilliseconds", 0.001),
    ] {
        let m = *mult;
        vm.register_host_fn("vybe:types", &format!("timeSpan{}", name), Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let val = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let total_secs = val * m;
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Arc::from("TimeSpan")));
            obj.properties.insert("totalseconds".into(), Value::F64(total_secs));
            obj.properties.insert("totalminutes".into(), Value::F64(total_secs / 60.0));
            obj.properties.insert("totalhours".into(), Value::F64(total_secs / 3600.0));
            obj.properties.insert("totaldays".into(), Value::F64(total_secs / 86400.0));
            obj.properties.insert("totalmilliseconds".into(), Value::F64(total_secs * 1000.0));
            let abs = total_secs.abs() as u64;
            obj.properties.insert("days".into(), Value::F64((abs / 86400) as f64));
            obj.properties.insert("hours".into(), Value::F64(((abs % 86400) / 3600) as f64));
            obj.properties.insert("minutes".into(), Value::F64(((abs % 3600) / 60) as f64));
            obj.properties.insert("seconds".into(), Value::F64((abs % 60) as f64));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));
    }
    vm.register_host_fn("vybe:types", "timeSpanZero", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("TimeSpan")));
        obj.properties.insert("totalseconds".into(), Value::F64(0.0));
        obj.properties.insert("totalminutes".into(), Value::F64(0.0));
        obj.properties.insert("totalhours".into(), Value::F64(0.0));
        obj.properties.insert("totaldays".into(), Value::F64(0.0));
        obj.properties.insert("totalmilliseconds".into(), Value::F64(0.0));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
}

// ============================================================
// Guid
// ============================================================

fn register_guid(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "guidNewGuid", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        // Simple UUID v4 using random bytes
        let mut bytes = [0u8; 16];
        let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
        let seed = t.as_nanos() as u64;
        // xorshift64 for pseudo-random
        let mut s = seed;
        for b in &mut bytes {
            s ^= s << 13; s ^= s >> 7; s ^= s << 17;
            *b = s as u8;
        }
        bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
        bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant 1
        let guid = format!(
            "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
            bytes[0], bytes[1], bytes[2], bytes[3],
            bytes[4], bytes[5], bytes[6], bytes[7],
            bytes[8], bytes[9], bytes[10], bytes[11],
            bytes[12], bytes[13], bytes[14], bytes[15]
        );
        Value::String(Arc::from(guid.as_str()))
    }));
    vm.register_host_fn("vybe:types", "guidEmpty", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::String(Arc::from("00000000-0000-0000-0000-000000000000"))
    }));
    vm.register_host_fn("vybe:types", "guidParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(s.as_str()))
    }));
}

// ============================================================
// Primitive type statics (Double, Single, Boolean, Decimal)
// ============================================================

fn register_primitives(vm: &mut VM) {
    // Double
    vm.register_host_fn("vybe:types", "doubleParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(f64::NAN))
    }));
    vm.register_host_fn("vybe:types", "doubleTryParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::Bool(s.trim().parse::<f64>().is_ok())
    }));

    // Boolean
    vm.register_host_fn("vybe:types", "booleanParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default().to_lowercase();
        Value::Bool(s == "true" || s == "1" || s == "yes")
    }));

    // Array static methods
    vm.register_host_fn("vybe:types", "arrayClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                for e in elems.iter_mut() { *e = Value::Null; }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arrayCopy", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(src)), Some(Value::Object(dst))) = (args.first(), args.get(1)) {
            let s = src.lock().unwrap();
            let mut d = dst.lock().unwrap();
            if let (ObjectKind::Array(src_elems), ObjectKind::Array(dst_elems)) = (&s.kind, &mut d.kind) {
                let count = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(src_elems.len());
                for i in 0..count.min(src_elems.len()).min(dst_elems.len()) {
                    dst_elems[i] = src_elems[i].clone();
                }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arrayResize", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let new_size = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.resize(new_size, Value::Null);
                o.properties.insert("length".into(), Value::F64(new_size as f64));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arraySort", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
            }
        }
        Value::Null
    }));
}
