fn register_before(vm: &mut VM) {
    array_fn(
        vm,
        "concat-before",
        vec![arr(), arr()],
        vec![arr()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    out.reserve(v.len());
                    out.extend(v.iter().cloned());
                }
            }
            // §23.1.3.1 Array.prototype.concat step 5.b — IsConcatSpreadable.
            //
            // ⛔ THE QUESTION IS NOT "IS IT ITERABLE". §22.1.3.1.1 says: a
            // non-Object is NEVER spread; otherwise ask `@@isConcatSpreadable`,
            // and only when that is `undefined` fall back to `IsArray`. This
            // used to spread by KIND, which got three things wrong at once —
            // a STRING was spread into its code points (`[].concat("ab")` gave
            // `a,b` where the spec appends `"ab"` whole), an array-like opting
            // IN with the symbol was appended as one element, and an array
            // opting OUT with `false` was spread anyway.
            match args.get(1) {
                Some(Value::Object(o)) => {
                    let (explicit, is_array, length) = {
                        let lock = o.lock().unwrap();
                        let explicit = lock.properties.get("isconcatspreadable").cloned();
                        let is_array = matches!(lock.kind, ObjectKind::Array(_));
                        let len = lock
                            .properties
                            .get("length")
                            .map(|v| v.as_f64() as usize)
                            .unwrap_or(0);
                        (explicit, is_array, len)
                    };
                    let spreadable = match explicit {
                        Some(Value::Undefined) | None => is_array,
                        // ⛔ ToBoolean (§7.1.2), NOT `Value::as_bool`. That
                        // accessor is `matches!(self, Bool(true))` — it answers
                        // "is this the boolean true", so every truthy NON-bool
                        // read as false. §22.1.3.1.1 step 4 applies ToBoolean,
                        // so `arr[@@isConcatSpreadable] = 1` must spread just
                        // as `= true` does; measured, `1` did not and `true`
                        // did.
                        Some(v) => crate::boolean::to_boolean(&v),
                    };
                    if !spreadable {
                        out.push(Value::Object(o.clone()));
                    } else {
                        let lock = o.lock().unwrap();
                        match &lock.kind {
                            ObjectKind::Array(v) => {
                                out.reserve(v.len());
                                out.extend(v.iter().cloned());
                            }
                            ObjectKind::Set(sv) => {
                                out.reserve(sv.len());
                                out.extend(sv.iter().cloned());
                            }
                            ObjectKind::Map(m) => {
                                out.reserve(m.len());
                                for (k, v) in m {
                                    out.push(crate::array::make_pair_array(k.clone(), v.clone()));
                                }
                            }
                            _ => {
                                // Array-LIKE that opted in: spread by `length`, reading
                                // the indexed properties (§23.1.3.1 step 5.c.iv).
                                out.reserve(length);
                                for i in 0..length {
                                    out.push(crate::keys::with_index_key(i, |key| {
                                        lock.properties
                                            .get(key)
                                            .cloned()
                                            .unwrap_or(Value::Undefined)
                                    }));
                                }
                            }
                        }
                    }
                }
                // A primitive is never spread — including a String, whose
                // `Symbol.iterator` is irrelevant here.
                Some(v) => out.push(v.clone()),
                None => {}
            }
            make_array(out)
        }),
    );
}
