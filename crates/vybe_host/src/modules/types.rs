//! Built-in .NET types: DateTime, StringBuilder, List, Dictionary.
//! Each constructor creates an object with methods as HostFunctions.

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
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
    vm.register_host_fn("vybe:types", "dateTimeNow", Box::new(|_vm: &mut VM, _args: &[Value]| {
        make_datetime_from_epoch_secs(epoch_secs())
    }));

    // DateTime.Parse(str) → parse a date string
    vm.register_host_fn("vybe:types", "dateTimeParse", Box::new(|_vm: &mut VM, args: &[Value]| {
        // Simplified: just store the string, parse on demand
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("DateTime")));
        obj.properties.insert("__raw".into(), Value::String(Rc::from(s.as_str())));
        obj.properties.insert("__epoch".into(), Value::F64(0.0)); // TODO: parse
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New DateTime(year, month, day) or New DateTime(year, month, day, hour, min, sec)
    vm.register_host_fn("vybe:types", "dateTimeNew", Box::new(|_vm: &mut VM, args: &[Value]| {
        let _this = args.first(); // ignore this from New
        let year = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(2000);
        let month = args.get(2).map(|v| v.as_f64() as u64).unwrap_or(1);
        let day = args.get(3).map(|v| v.as_f64() as u64).unwrap_or(1);
        let hour = args.get(4).map(|v| v.as_f64() as u64).unwrap_or(0);
        let min = args.get(5).map(|v| v.as_f64() as u64).unwrap_or(0);
        let sec = args.get(6).map(|v| v.as_f64() as u64).unwrap_or(0);

        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("DateTime")));
        obj.properties.insert("year".into(), Value::F64(year as f64));
        obj.properties.insert("month".into(), Value::F64(month as f64));
        obj.properties.insert("day".into(), Value::F64(day as f64));
        obj.properties.insert("hour".into(), Value::F64(hour as f64));
        obj.properties.insert("minute".into(), Value::F64(min as f64));
        obj.properties.insert("second".into(), Value::F64(sec as f64));

        let epoch = date_to_epoch(year, month, day, hour, min, sec);
        obj.properties.insert("__epoch".into(), Value::F64(epoch as f64));

        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Instance methods called via vybe:runtime/callMethod or directly
    vm.register_host_fn("vybe:types", "dateTimeAddDays", Box::new(|_vm: &mut VM, args: &[Value]| {
        dt_add(args, 86400.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddHours", Box::new(|_vm: &mut VM, args: &[Value]| {
        dt_add(args, 3600.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddMinutes", Box::new(|_vm: &mut VM, args: &[Value]| {
        dt_add(args, 60.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddSeconds", Box::new(|_vm: &mut VM, args: &[Value]| {
        dt_add(args, 1.0)
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddMonths", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let months = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(0);
            let (y, m, d, h, min, s) = decompose(epoch);
            let total_months = y * 12 + m as i64 + months;
            let ny = total_months / 12;
            let nm = ((total_months % 12) + 12) % 12;
            let nm = if nm == 0 { 12 } else { nm as u64 };
            return make_datetime_from_epoch_secs(date_to_epoch(ny, nm, d, h, min, s));
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "dateTimeAddYears", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let years = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(0);
            let (y, m, d, h, min, s) = decompose(epoch);
            return make_datetime_from_epoch_secs(date_to_epoch(y + years, m, d, h, min, s));
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "dateTimeToString", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let (y, m, d, h, min, s) = decompose(epoch);
            return Value::String(Rc::from(format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, m, d, h, min, s).as_str()));
        }
        Value::String(Rc::from(""))
    }));
    vm.register_host_fn("vybe:types", "dateTimeToShortDate", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            let (y, m, d, _, _, _) = decompose(epoch);
            return Value::String(Rc::from(format!("{:02}/{:02}/{:04}", m, d, y).as_str()));
        }
        Value::String(Rc::from(""))
    }));
}

// ============================================================
// StringBuilder
// ============================================================

fn register_stringbuilder(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "stringBuilderNew", Box::new(|_vm: &mut VM, args: &[Value]| {
        let _this = args.first();
        let initial = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("StringBuilder")));
        obj.properties.insert("__buffer".into(), Value::String(Rc::from(initial.as_str())));
        obj.properties.insert("length".into(), Value::F64(initial.len() as f64));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "sbAppend", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let new_buf = format!("{}{}", current, text);
            let len = new_buf.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Rc::from(new_buf.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null) // return this for chaining
    }));

    vm.register_host_fn("vybe:types", "sbAppendLine", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let new_buf = format!("{}{}\n", current, text);
            let len = new_buf.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Rc::from(new_buf.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbToString", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            return o.properties.get("__buffer").cloned().unwrap_or(Value::String(Rc::from("")));
        }
        Value::String(Rc::from(""))
    }));

    vm.register_host_fn("vybe:types", "sbClear", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            o.properties.insert("__buffer".into(), Value::String(Rc::from("")));
            o.properties.insert("length".into(), Value::F64(0.0));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbInsert", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let text = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
            let mut current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let idx = index.min(current.len());
            current.insert_str(idx, &text);
            let len = current.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Rc::from(current.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("vybe:types", "sbReplace", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let old = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let new = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
            let current = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let result = current.replace(&old, &new);
            let len = result.len() as f64;
            o.properties.insert("__buffer".into(), Value::String(Rc::from(result.as_str())));
            o.properties.insert("length".into(), Value::F64(len));
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

// ============================================================
// List(Of T) — backed by array
// ============================================================

fn register_list(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "listNew", Box::new(|_vm: &mut VM, args: &[Value]| {
        let _this = args.first();
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Rc::from("List")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // List.Add(item)
    vm.register_host_fn("vybe:types", "listAdd", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "listRemove", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "listRemoveAt", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "listContains", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search));
            }
        }
        Value::Bool(false)
    }));

    // List.Count → number
    vm.register_host_fn("vybe:types", "listCount", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::F64(elems.len() as f64);
            }
        }
        Value::F64(0.0)
    }));

    // List.Clear()
    vm.register_host_fn("vybe:types", "listClear", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.clear();
                o.properties.insert("length".into(), Value::F64(0.0));
                o.properties.insert("count".into(), Value::F64(0.0));
            }
        }
        Value::Null
    }));

    // List.IndexOf(item) → index or -1
    vm.register_host_fn("vybe:types", "listIndexOf", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                for (i, e) in elems.iter().enumerate() {
                    if format!("{}", e) == search { return Value::F64(i as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));

    // List.Item(index) → element at index
    vm.register_host_fn("vybe:types", "listItem", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.get(idx).cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // List.Insert(index, item)
    vm.register_host_fn("vybe:types", "listInsert", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let item = args.get(2).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "listAddRange", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(1)) {
            let s = src.borrow();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items = src_elems.clone();
                drop(s);
                let mut d = dst.borrow_mut();
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
    vm.register_host_fn("vybe:types", "listSort", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
            }
        }
        Value::Null
    }));

    // List.Reverse()
    vm.register_host_fn("vybe:types", "listReverse", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.reverse();
            }
        }
        Value::Null
    }));

    // List.ToArray() → new array copy
    vm.register_host_fn("vybe:types", "listToArray", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Object(Rc::new(RefCell::new(Object::new_array(elems.clone()))));
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));
}

// ============================================================
// Dictionary(Of K, V)
// ============================================================

fn register_dictionary(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "dictNew", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Dictionary")));
        obj.properties.insert("__data".into(), Value::Object(Rc::new(RefCell::new(Object::new()))));
        obj.properties.insert("count".into(), Value::F64(0.0));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Dict.Add(key, value) / Dict.Item(key) = value
    vm.register_host_fn("vybe:types", "dictAdd", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                data.borrow_mut().properties.insert(key, value);
                let count = data.borrow().properties.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("count".into(), Value::F64(count));
            }
        }
        Value::Null
    }));

    // Dict.Item(key) → value
    vm.register_host_fn("vybe:types", "dictItem", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return data.borrow().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            // Direct property lookup (struct_new-based dicts)
            if let Some(val) = o.properties.get(&key) {
                return val.clone();
            }
        }
        Value::Null
    }));

    // Dict.ContainsKey(key) → bool
    vm.register_host_fn("vybe:types", "dictContainsKey", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return Value::Bool(data.borrow().properties.contains_key(&key));
            }
        }
        Value::Bool(false)
    }));

    // Dict.Remove(key) → bool
    vm.register_host_fn("vybe:types", "dictRemove", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let removed = data.borrow_mut().properties.remove(&key).is_some();
                let count = data.borrow().properties.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("count".into(), Value::F64(count));
                return Value::Bool(removed);
            }
        }
        Value::Bool(false)
    }));

    // Dict.Keys → array
    vm.register_host_fn("vybe:types", "dictKeys", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            // Try __data first (old-style dicts), then enumerate properties directly
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let keys: Vec<Value> = data.borrow().properties.keys()
                    .map(|k| Value::String(Rc::from(k.as_str())))
                    .collect();
                return Value::Object(Rc::new(RefCell::new(Object::new_array(keys))));
            }
            // Direct property enumeration (struct_new-based dicts)
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| !k.starts_with("__")) // skip internal properties
                .map(|k| Value::String(Rc::from(k.as_str())))
                .collect();
            if !keys.is_empty() {
                return Value::Object(Rc::new(RefCell::new(Object::new_array(keys))));
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Dict.Values → array
    vm.register_host_fn("vybe:types", "dictValues", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let vals: Vec<Value> = data.borrow().properties.values().cloned().collect();
                return Value::Object(Rc::new(RefCell::new(Object::new_array(vals))));
            }
            let vals: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(_, v)| v.clone())
                .collect();
            if !vals.is_empty() {
                return Value::Object(Rc::new(RefCell::new(Object::new_array(vals))));
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Dict.Clear()
    vm.register_host_fn("vybe:types", "dictClear", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                data.borrow_mut().properties.clear();
                drop(o);
                obj.borrow_mut().properties.insert("count".into(), Value::F64(0.0));
            }
        }
        Value::Null
    }));
}

// ============================================================
// Process
// ============================================================

fn register_process(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "processStart", Box::new(|_vm: &mut VM, args: &[Value]| {
        let cmd = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let cmd_args = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        match std::process::Command::new(&cmd).args(cmd_args.split_whitespace()).output() {
            Ok(output) => {
                let stdout = String::from_utf8_lossy(&output.stdout);
                Value::String(Rc::from(stdout.as_ref()))
            }
            Err(e) => Value::String(Rc::from(format!("Error: {}", e).as_str())),
        }
    }));
}

// ============================================================
// DateTime helpers
// ============================================================

fn epoch_secs() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
}

fn make_datetime_from_epoch_secs(secs: u64) -> Value {
    let (y, m, d, h, min, s) = decompose(secs);
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Rc::from("DateTime")));
    obj.properties.insert("__epoch".into(), Value::F64(secs as f64));
    obj.properties.insert("year".into(), Value::F64(y as f64));
    obj.properties.insert("month".into(), Value::F64(m as f64));
    obj.properties.insert("day".into(), Value::F64(d as f64));
    obj.properties.insert("hour".into(), Value::F64(h as f64));
    obj.properties.insert("minute".into(), Value::F64(min as f64));
    obj.properties.insert("second".into(), Value::F64(s as f64));
    Value::Object(Rc::new(RefCell::new(obj)))
}

fn dt_add(args: &[Value], multiplier: f64) -> Value {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.borrow();
        let epoch = o.properties.get("__epoch").map(|v| v.as_f64()).unwrap_or(0.0);
        let amount = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let new_epoch = (epoch + amount * multiplier) as u64;
        return make_datetime_from_epoch_secs(new_epoch);
    }
    Value::Null
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

// ============================================================
// Queue / Stack / HashSet
// ============================================================

fn register_queue_stack(vm: &mut VM) {
    // Queue — backed by array, FIFO
    vm.register_host_fn("vybe:types", "queueNew", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Rc::from("Queue")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "queueEnqueue", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.push(item);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "queueDequeue", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "queuePeek", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.first().cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // Stack — backed by array, LIFO
    vm.register_host_fn("vybe:types", "stackNew", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Rc::from("Stack")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "stackPush", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.push(item);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "stackPop", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let val = elems.pop().unwrap_or(Value::Null);
                let len = elems.len() as f64;
                o.properties.insert("count".into(), Value::F64(len));
                return val;
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "stackPeek", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return elems.last().cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // HashSet — backed by array with uniqueness
    vm.register_host_fn("vybe:types", "hashSetNew", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Rc::from("HashSet")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));
    vm.register_host_fn("vybe:types", "hashSetAdd", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let item_str = format!("{}", item);
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:types", "hashSetContains", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search));
            }
        }
        Value::Bool(false)
    }));
    vm.register_host_fn("vybe:types", "hashSetRemove", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = obj.borrow_mut();
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
        vm.register_host_fn("vybe:types", &format!("timeSpan{}", name), Box::new(move |_vm: &mut VM, args: &[Value]| {
            let val = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let total_secs = val * m;
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Rc::from("TimeSpan")));
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
            Value::Object(Rc::new(RefCell::new(obj)))
        }));
    }
    vm.register_host_fn("vybe:types", "timeSpanZero", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("TimeSpan")));
        obj.properties.insert("totalseconds".into(), Value::F64(0.0));
        obj.properties.insert("totalminutes".into(), Value::F64(0.0));
        obj.properties.insert("totalhours".into(), Value::F64(0.0));
        obj.properties.insert("totaldays".into(), Value::F64(0.0));
        obj.properties.insert("totalmilliseconds".into(), Value::F64(0.0));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));
}

// ============================================================
// Guid
// ============================================================

fn register_guid(vm: &mut VM) {
    vm.register_host_fn("vybe:types", "guidNewGuid", Box::new(|_vm: &mut VM, _args: &[Value]| {
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
        Value::String(Rc::from(guid.as_str()))
    }));
    vm.register_host_fn("vybe:types", "guidEmpty", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::String(Rc::from("00000000-0000-0000-0000-000000000000"))
    }));
    vm.register_host_fn("vybe:types", "guidParse", Box::new(|_vm: &mut VM, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Rc::from(s.as_str()))
    }));
}

// ============================================================
// Primitive type statics (Double, Single, Boolean, Decimal)
// ============================================================

fn register_primitives(vm: &mut VM) {
    // Double
    vm.register_host_fn("vybe:types", "doubleParse", Box::new(|_vm: &mut VM, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(f64::NAN))
    }));
    vm.register_host_fn("vybe:types", "doubleTryParse", Box::new(|_vm: &mut VM, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::Bool(s.trim().parse::<f64>().is_ok())
    }));

    // Boolean
    vm.register_host_fn("vybe:types", "booleanParse", Box::new(|_vm: &mut VM, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default().to_lowercase();
        Value::Bool(s == "true" || s == "1" || s == "yes")
    }));

    // Array static methods
    vm.register_host_fn("vybe:types", "arrayClear", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                for e in elems.iter_mut() { *e = Value::Null; }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arrayCopy", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let (Some(Value::Object(src)), Some(Value::Object(dst))) = (args.first(), args.get(1)) {
            let s = src.borrow();
            let mut d = dst.borrow_mut();
            if let (ObjectKind::Array(src_elems), ObjectKind::Array(dst_elems)) = (&s.kind, &mut d.kind) {
                let count = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(src_elems.len());
                for i in 0..count.min(src_elems.len()).min(dst_elems.len()) {
                    dst_elems[i] = src_elems[i].clone();
                }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arrayResize", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let new_size = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.resize(new_size, Value::Null);
                o.properties.insert("length".into(), Value::F64(new_size as f64));
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:types", "arraySort", Box::new(|_vm: &mut VM, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(elems) = &mut o.kind {
                elems.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
            }
        }
        Value::Null
    }));
}
