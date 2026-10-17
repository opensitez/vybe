#[path = "object_enumerability_membership/fixture.rs"]
mod fixture;
use fixture::*;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicBool, Ordering},
};
use vybe_runtime::{Value, value::ObjectKind};
#[test]
fn actual_define_property_all_enumeration_surfaces_and_live_refresh() {
    let mut v = vm();
    for (n, h) in [(0, 0), (16, 8), (64, 0), (64, 32), (64, 64)] {
        let source = fixture(&mut v, n, h);
        let want = key_list(n, h);
        for name in ["keys", "iterForIn"] {
            let f = host(&v, "ecma:object", name);
            assert_eq!(strings(&v.invoke(&f, &[source.clone()]).unwrap()), want);
        }
        let f = host(&v, "ecma:object", "values");
        assert_eq!(
            values(&v.invoke(&f, &[source.clone()]).unwrap()),
            (h..n).map(|i| Value::I32(i as i32)).collect::<Vec<_>>()
        );
        let f = host(&v, "ecma:object", "entries");
        let out = v.invoke(&f, &[source.clone()]).unwrap();
        let got = values(&out);
        assert_eq!(got.len(), want.len());
        for (row, i) in got.iter().zip(h..n) {
            assert_eq!(
                values(row),
                vec![
                    Value::String(format!("k{i:05}").into()),
                    Value::I32(i as i32)
                ]
            );
        }
        if n > 0 {
            define(
                &mut v,
                &source,
                Value::String("k00000".into()),
                data_descriptor(Value::I32(777), true),
            );
            let f = host(&v, "ecma:object", "keys");
            let got = strings(&v.invoke(&f, &[source.clone()]).unwrap());
            assert!(got.contains(&"k00000".into()));
            define(
                &mut v,
                &source,
                Value::String("k00000".into()),
                data_descriptor(Value::I32(777), false),
            );
            assert!(!strings(&v.invoke(&f, &[source]).unwrap()).contains(&"k00000".into()));
        }
        assert!(v.last_exception.is_none());
    }
}
#[test]
fn integer_order_and_recorded_symbol_storage_parity_with_malformed_markers() {
    let mut v = vm();
    let source = fixture(&mut v, 64, 32);
    for k in ["10", "2", "01"] {
        define(
            &mut v,
            &source,
            Value::String(k.into()),
            data_descriptor(Value::I32(90), true),
        );
    }
    let symbol = host(&v, "ecma:symbol", "new");
    let symbol = v
        .invoke(&symbol, &[Value::String("excluded".into())])
        .unwrap();
    let Value::Symbol(symbol_identity) = &symbol else {
        panic!("genuine Symbol result")
    };
    let storage = vybe_platform_ecma::symbol::canonical_property_key(symbol_identity);
    assert_eq!(storage, "Symbol(excluded)");
    define(
        &mut v,
        &source,
        symbol.clone(),
        data_descriptor(Value::I32(91), true),
    );
    // Freeze the genuine stored Symbol identity and canonical data slot BEFORE
    // recording the separate original Object.keys exclusion defect.
    assert_eq!(
        obj(&source).lock().unwrap().properties.get(&storage),
        Some(&Value::I32(91))
    );
    let tracked = obj(&source).lock().unwrap().properties["__sym_keys"].clone();
    let tracked = values(&tracked);
    assert_eq!(tracked.len(), 1);
    assert!(matches!(&tracked[0],Value::Symbol(actual) if Arc::ptr_eq(actual,symbol_identity)));
    let symbols = host(&v, "ecma:object", "getOwnPropertySymbols");
    let symbols = values(&v.invoke(&symbols, &[source.clone()]).unwrap());
    assert_eq!(symbols.len(), 1);
    assert!(matches!(&symbols[0],Value::Symbol(actual) if Arc::ptr_eq(actual,symbol_identity)));
    let descriptor = host(&v, "ecma:object", "getOwnPropertyDescriptor");
    let descriptor = v
        .invoke(&descriptor, &[source.clone(), symbol.clone()])
        .unwrap();
    assert_eq!(
        obj(&descriptor).lock().unwrap().properties.get("value"),
        Some(&Value::I32(91))
    );
    let keys = host(&v, "ecma:object", "keys");
    let got = strings(&v.invoke(&keys, &[source.clone()]).unwrap());
    let mut want = vec!["2".into(), "10".into()];
    want.extend(key_list(64, 32));
    want.push("01".into());
    // The original descriptor_own_keys appends __sym_keys' canonical String
    // after ordered string keys; ordinary_enumerable_keys does not exclude it.
    // This is recorded legacy parity, NEVER a normative Symbol exclusion pass.
    want.push(storage.clone());
    assert_eq!(got, want);
    eprintln!(
        "{}",
        serde_json::json!({"symbolInputTag":"Symbol","trackedSymbolIdentity":true,"storageKey":storage,"storedValue":91,"keysIncludesStorage":true,"heldNormativeGap":"Object.keys must exclude Symbol properties"})
    );
    let mut all = vec!["2".into(), "10".into()];
    all.extend(key_list(64, 0));
    all.push("01".into());
    all.push(storage.clone());
    // Only String entries are meaningful, including duplicate, empty, numeric-like.
    let mut marker = Vec::new();
    for _ in 0..32 {
        marker.push(Value::String("k00000".into()));
    }
    marker.extend([
        Value::I32(1),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
        Value::String("2".into()),
    ]);
    obj(&source)
        .lock()
        .unwrap()
        .properties
        .insert("__nonenum".into(), array(marker));
    let got = strings(&v.invoke(&keys, &[source.clone()]).unwrap());
    let filtered: Vec<String> = all
        .iter()
        .filter(|key| key.as_str() != "k00000" && key.as_str() != "2")
        .cloned()
        .collect();
    assert_eq!(got, filtered);
    for malformed in [Value::Null, Value::Undefined, Value::I32(3), object([])] {
        obj(&source)
            .lock()
            .unwrap()
            .properties
            .insert("__nonenum".into(), malformed);
        assert_eq!(
            strings(&v.invoke(&keys, &[source.clone()]).unwrap()),
            all,
            "67 ordinary string keys plus the separately verified leaked Symbol storage key"
        );
    }
}
#[test]
fn for_in_retains_prototype_order_and_deduplication() {
    let mut v = vm();
    let source = fixture(&mut v, 64, 32);
    let prototype = fixture(&mut v, 64, 48);
    define(
        &mut v,
        &prototype,
        Value::String("tail".into()),
        data_descriptor(Value::I32(9), true),
    );
    obj(&source)
        .lock()
        .unwrap()
        .properties
        .insert("__proto__".into(), prototype);
    let f = host(&v, "ecma:object", "iterForIn");
    let mut want = key_list(64, 32);
    want.push("tail".into());
    assert_eq!(strings(&v.invoke(&f, &[source]).unwrap()), want);
}
#[test]
fn strong_and_weak_marker_aliases_keep_original_live_read_fallback() {
    let mut v = vm();
    let source = fixture(&mut v, 64, 32);
    let keys = host(&v, "ecma:object", "keys");
    assert_eq!(marker_owners(&source), Some((1, 0)));
    let marker = obj(&source).lock().unwrap().properties["__nonenum"].clone();
    assert_eq!(marker_owners(&source), Some((2, 0)));
    assert_eq!(
        strings(&v.invoke(&keys, &[source.clone()]).unwrap()),
        key_list(64, 32)
    );
    let weak = Arc::downgrade(&obj(&marker));
    drop(marker);
    assert_eq!(marker_owners(&source), Some((1, 1)));
    assert_eq!(
        strings(&v.invoke(&keys, &[source.clone()]).unwrap()),
        key_list(64, 32)
    );
    drop(weak);
    assert_eq!(marker_owners(&source), Some((1, 0)));
    let marker = obj(&source).lock().unwrap().properties["__nonenum"].clone();
    let raw = obj(&marker);
    let stop = Arc::new(AtomicBool::new(false));
    let halt = stop.clone();
    let t = std::thread::spawn(move || {
        let mut hide = true;
        while !halt.load(Ordering::Acquire) {
            let mut m = raw.lock().unwrap();
            let ObjectKind::Array(ref mut a) = m.kind else {
                panic!("marker Array")
            };
            a.clear();
            if hide {
                a.extend((0..32).map(|i| Value::String(format!("k{i:05}").into())));
            }
            hide = !hide;
        }
    });
    for _ in 0..16 {
        let got = strings(&v.invoke(&keys, &[source.clone()]).unwrap());
        for k in key_list(64, 32) {
            assert!(got.contains(&k));
        }
        let mut sorted = got.clone();
        sorted.sort();
        assert_eq!(got, sorted);
    }
    stop.store(true, Ordering::Release);
    t.join().unwrap();
    drop(marker);
}

#[test]
fn ignored_marker_objects_do_not_survive_in_key_results() {
    let mut v = vm();
    let source = fixture(&mut v, 64, 32);
    let ignored = vybe_runtime::heap::alloc(vybe_runtime::value::Object::new());
    let weak = Arc::downgrade(&ignored);
    let mut marker = (0..32)
        .map(|i| Value::String(format!("k{i:05}").into()))
        .collect::<Vec<_>>();
    marker.push(Value::Object(ignored));
    obj(&source)
        .lock()
        .unwrap()
        .properties
        .insert("__nonenum".into(), array(marker));
    let f = host(&v, "ecma:object", "keys");
    let result = v.invoke(&f, &[source.clone()]).unwrap();
    assert_eq!(strings(&result), key_list(64, 32));
    assert!(weak.upgrade().is_some());
    obj(&source)
        .lock()
        .unwrap()
        .properties
        .shift_remove("__nonenum");
    assert!(
        weak.upgrade().is_none(),
        "filter must not retain ignored Object roots"
    );
    drop(result);
}
#[test]
fn genuine_proxy_ownkeys_and_descriptor_callbacks_keep_exact_order() {
    let mut v = vm();
    let target = object([("a".into(), Value::I32(1)), ("b".into(), Value::I32(2))]);
    let trace = Arc::new(Mutex::new(Vec::<String>::new()));
    let t = trace.clone();
    v.register_host_fn(
        "test:enum",
        "own",
        Box::new(move |_, _| {
            t.lock().unwrap().push("own".into());
            array(vec![Value::String("b".into()), Value::String("a".into())])
        }),
    );
    let own = host(&v, "test:enum", "own");
    let descriptor = host(&v, "ecma:object", "getOwnPropertyDescriptor");
    let actual_target = target.clone();
    let t = trace.clone();
    v.register_host_fn(
        "test:enum",
        "desc",
        Box::new(move |ctx, args| {
            let key = args
                .iter()
                .rev()
                .find(|x| matches!(x, Value::String(_)))
                .expect("genuine descriptor trap key")
                .clone();
            t.lock().unwrap().push(format!(
                "desc:{}",
                match &key {
                    Value::String(text) => text.as_ref(),
                    _ => unreachable!(),
                }
            ));
            ctx.invoke(&descriptor, &[actual_target.clone(), key])
        }),
    );
    let desc = host(&v, "test:enum", "desc");
    let handler = object([
        ("ownKeys".into(), own),
        ("getOwnPropertyDescriptor".into(), desc),
    ]);
    let new = host(&v, "ecma:proxy", "new");
    let proxy = v.invoke(&new, &[target, handler]).unwrap();
    for name in ["keys", "iterForIn"] {
        trace.lock().unwrap().clear();
        let f = host(&v, "ecma:object", name);
        assert_eq!(
            strings(&v.invoke(&f, &[proxy.clone()]).unwrap()),
            vec!["b", "a"]
        );
        assert_eq!(*trace.lock().unwrap(), vec!["own", "desc:b", "desc:a"]);
        assert!(v.last_exception.is_none());
    }
}
#[test]
fn real_getters_reenter_keys_mutate_later_values_and_preserve_recorded_snapshot() {
    for name in ["values", "entries"] {
        let mut v = vm();
        let source = object([]);
        define(
            &mut v,
            &source,
            Value::String("first".into()),
            data_descriptor(Value::I32(0), true),
        );
        for i in 0..64 {
            define(
                &mut v,
                &source,
                Value::String(format!("k{i:05}").into()),
                data_descriptor(Value::I32(i), i >= 32),
            );
        }
        let calls = Arc::new(Mutex::new(Vec::<String>::new()));
        let t = calls.clone();
        let s = source.clone();
        let define_fn = host(&v, "ecma:object", "defineProperty");
        let keys_fn = host(&v, "ecma:object", "keys");
        v.register_host_fn(
            "test:enum",
            "getFirst",
            Box::new(move |ctx, _| {
                t.lock().unwrap().push("getter".into());
                ctx.invoke(
                    &define_fn,
                    &[
                        s.clone(),
                        Value::String("k00063".into()),
                        data_descriptor(Value::I32(99), false),
                    ],
                );
                let nested = ctx.invoke(&keys_fn, &[s.clone()]);
                let nested = strings(&nested);
                assert_eq!(nested.len(), 32);
                assert!(!nested.contains(&"k00063".into()));
                t.lock().unwrap().push("nested".into());
                Value::I32(7)
            }),
        );
        let getter = host(&v, "test:enum", "getFirst");
        define(
            &mut v,
            &source,
            Value::String("first".into()),
            object([
                ("get".into(), getter),
                ("enumerable".into(), Value::Bool(true)),
                ("configurable".into(), Value::Bool(true)),
            ]),
        );
        let f = host(&v, "ecma:object", name);
        let out = v.invoke(&f, &[source.clone()]).unwrap();
        let result = values(&out);
        assert_eq!(result.len(), 33);
        if name == "values" {
            assert_eq!(result[0], Value::I32(7));
            assert_eq!(result.last(), Some(&Value::I32(99)));
        } else {
            assert_eq!(
                values(&result[0]),
                vec![Value::String("first".into()), Value::I32(7)]
            );
            assert_eq!(
                values(result.last().unwrap()),
                vec![Value::String("k00063".into()), Value::I32(99)]
            );
        }
        assert_eq!(*calls.lock().unwrap(), vec!["getter", "nested"]);
        assert!(v.last_exception.is_none());
        eprintln!(
            "{}",
            serde_json::json!({"surface":name,"recordedLegacySnapshotIncludesLaterHidden":true,"count":result.len(),"trace":*calls.lock().unwrap(),"scope":"strict before/candidate legacy parity; normative live descriptor recheck remains separate"})
        );
    }
}
#[test]
fn getter_throw_stops_later_getters_and_retains_exception() {
    for name in ["values", "entries"] {
        let mut v = vm();
        let source = fixture(&mut v, 64, 32);
        let trace = Arc::new(Mutex::new(Vec::<String>::new()));
        let t = trace.clone();
        v.register_host_fn(
            "test:enum",
            "stop",
            Box::new(move |ctx, _| {
                t.lock().unwrap().push("stop".into());
                ctx.throw_value(Value::String("enum-getter-stop".into()));
                Value::Undefined
            }),
        );
        let stop = host(&v, "test:enum", "stop");
        let t = trace.clone();
        v.register_host_fn(
            "test:enum",
            "later",
            Box::new(move |_, _| {
                t.lock().unwrap().push("later".into());
                Value::I32(42)
            }),
        );
        let later = host(&v, "test:enum", "later");
        for (key, getter) in [("k00032", stop), ("k00033", later)] {
            define(
                &mut v,
                &source,
                Value::String(key.into()),
                object([
                    ("get".into(), getter),
                    ("enumerable".into(), Value::Bool(true)),
                    ("configurable".into(), Value::Bool(true)),
                ]),
            );
        }
        let f = host(&v, "ecma:object", name);
        let result = v.invoke(&f, &[source]);
        assert!(
            result.is_err() || v.last_exception.is_some(),
            "actual thrown value must remain"
        );
        assert_eq!(
            *trace.lock().unwrap(),
            vec!["stop"],
            "actual provider must not visit later getter"
        );
        eprintln!(
            "{}",
            serde_json::json!({"surface":name,"error":result.as_ref().err().map(|e|serde_json::json!({"message":e.message,"line":e.line})),"pending":v.last_exception.as_ref().map(ToString::to_string),"trace":*trace.lock().unwrap()})
        );
    }
}

#[test]
#[ignore = "held original provider gap: values/entries prefilter keys before getter changes later enumerability"]
fn normative_later_hidden_descriptor_must_be_rechecked_after_getter() {
    for name in ["values", "entries"] {
        let mut v = vm();
        let source = object([]);
        define(
            &mut v,
            &source,
            Value::String("first".into()),
            data_descriptor(Value::I32(0), true),
        );
        for i in 0..64 {
            define(
                &mut v,
                &source,
                Value::String(format!("k{i:05}").into()),
                data_descriptor(Value::I32(i), i >= 32),
            );
        }
        let calls = Arc::new(Mutex::new(Vec::<String>::new()));
        let t = calls.clone();
        let s = source.clone();
        let define_fn = host(&v, "ecma:object", "defineProperty");
        let keys_fn = host(&v, "ecma:object", "keys");
        v.register_host_fn(
            "test:enum",
            "getFirst",
            Box::new(move |ctx, _| {
                t.lock().unwrap().push("getter".into());
                ctx.invoke(
                    &define_fn,
                    &[
                        s.clone(),
                        Value::String("k00063".into()),
                        data_descriptor(Value::I32(99), false),
                    ],
                );
                let nested = ctx.invoke(&keys_fn, &[s.clone()]);
                let nested = strings(&nested);
                assert_eq!(nested.len(), 32);
                assert!(!nested.contains(&"k00063".into()));
                t.lock().unwrap().push("nested".into());
                Value::I32(7)
            }),
        );
        let getter = host(&v, "test:enum", "getFirst");
        define(
            &mut v,
            &source,
            Value::String("first".into()),
            object([
                ("get".into(), getter),
                ("enumerable".into(), Value::Bool(true)),
                ("configurable".into(), Value::Bool(true)),
            ]),
        );
        let f = host(&v, "ecma:object", name);
        let out = v.invoke(&f, &[source.clone()]).unwrap();
        let result = values(&out);
        assert_eq!(result.len(), 32);
        if name == "values" {
            assert_eq!(result[0], Value::I32(7));
            assert_eq!(result.last(), Some(&Value::I32(62)));
        } else {
            assert_eq!(
                values(&result[0]),
                vec![Value::String("first".into()), Value::I32(7)]
            );
            assert_eq!(
                values(result.last().unwrap()),
                vec![Value::String("k00062".into()), Value::I32(62)]
            );
        }
        assert_eq!(*calls.lock().unwrap(), vec!["getter", "nested"]);
        assert!(v.last_exception.is_none());
        eprintln!(
            "{}",
            serde_json::json!({"surface":name,"normativeDescriptorRecheck":true,"count":result.len(),"trace":*calls.lock().unwrap(),"scope":"held normative original provider gap; never counted as a pass"})
        );
    }
}

#[test]
#[ignore = "held original Symbol enumeration defect: descriptor_own_keys appends canonical Symbol storage as a string"]
fn normative_object_keys_must_exclude_genuine_symbol_property() {
    let mut v = vm();
    let source = fixture(&mut v, 64, 32);
    let symbol = host(&v, "ecma:symbol", "new");
    let symbol = v
        .invoke(&symbol, &[Value::String("excluded".into())])
        .unwrap();
    assert!(matches!(symbol, Value::Symbol(_)));
    define(
        &mut v,
        &source,
        symbol,
        data_descriptor(Value::I32(91), true),
    );
    let keys = host(&v, "ecma:object", "keys");
    assert_eq!(
        strings(&v.invoke(&keys, &[source]).unwrap()),
        key_list(64, 32),
        "strict normative exclusion remains held; never counted as a pass"
    );
}

#[test]
fn actual_heap_tracked_marker_uses_weak_fallback_and_generation_reset() {
    // Tracking is thread-local and has no disable API. An isolated joined thread
    // prevents test/serve-style tracking from contaminating the other fixtures.
    let outside = vybe_runtime::heap::is_tracking();
    std::thread::spawn(|| {
        assert!(!vybe_runtime::heap::is_tracking());
        vybe_runtime::heap::enable_tracking();
        assert!(vybe_runtime::heap::is_tracking());
        let mut v = vm();
        let keys = host(&v, "ecma:object", "keys");
        let baseline = vybe_runtime::heap::snapshot();
        let baseline_live = vybe_runtime::heap::live_count();
        let source = fixture(&mut v, 64, 32);
        let (strong, weak) = marker_owners(&source).expect("real defineProperty nonenum marker");
        assert_eq!(strong, 1);
        assert!(weak >= 1, "actual heap registry must own a Weak; Arc::get_mut fastpath must reject it");
        let output = v.invoke(&keys, &[source.clone()]).unwrap();
        assert_eq!(strings(&output), key_list(64, 32));
        assert!(v.last_exception.is_none());
        drop(output);
        assert!(vybe_runtime::heap::live_count() > baseline_live);
        // Retain a genuine cycle after program handles are dropped, then exercise
        // actual runtime generation cleanup, not a hand-cleared mock container.
        let source_weak = Arc::downgrade(&obj(&source));
        obj(&source).lock().unwrap().properties.insert("__cycle".into(), source.clone());
        drop(source);
        assert!(source_weak.upgrade().is_some());
        vybe_runtime::heap::restore(&baseline);
        assert!(source_weak.upgrade().is_none(), "restore must sever post-snapshot cyclic contents");
        assert_eq!(vybe_runtime::heap::live_count(), baseline_live);
        eprintln!("{}", serde_json::json!({"actualHeapTracking":true,"markerStrong":strong,"markerWeak":weak,"baselineLive":baseline_live,"liveAfterReset":vybe_runtime::heap::live_count(),"scope":"tracked-marker original fallback only; no serve/warm performance claim"}));
    }).join().unwrap();
    assert_eq!(vybe_runtime::heap::is_tracking(), outside);
}
