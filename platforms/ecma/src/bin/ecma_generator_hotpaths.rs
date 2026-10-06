//! Exercise the production helper without exposing benchmark-only library APIs.
mod implementation {
    include!("../generator.rs");

    fn original_next(genobj: &Arc<Mutex<Object>>) -> Value {
        let mut object = genobj.lock().unwrap();
        if matches!(object.properties.get(GEN_DONE), Some(Value::Bool(true))) {
            return iter_result(Value::Undefined, true);
        }
        let position = match object.properties.get(GEN_POS) {
            Some(Value::I32(position)) => *position as usize,
            _ => 0,
        };
        let items = match object.properties.get(GEN_ITEMS).cloned() {
            Some(Value::Object(items)) => items,
            _ => return iter_result(Value::Undefined, true),
        };
        let items = items.lock().unwrap();
        let snapshot = match &items.kind {
            ObjectKind::Array(items) => items.clone(),
            _ => return iter_result(Value::Undefined, true),
        };
        drop(items);
        if position >= snapshot.len() {
            object.properties.insert(GEN_DONE.into(), Value::Bool(true));
            return iter_result(Value::Undefined, true);
        }
        let value = snapshot[position].clone();
        object.properties.insert(GEN_POS.into(), Value::I32((position + 1) as i32));
        iter_result(value, false)
    }

    fn generator(items: Vec<Value>) -> Arc<Mutex<Object>> {
        match new_generator(items) {
            Value::Object(object) => object,
            _ => unreachable!(),
        }
    }

    fn result(value: Value) -> (Value, bool) {
        let Value::Object(object) = value else { panic!("iterator result object missing") };
        let object = object.lock().unwrap();
        (object.properties.get("value").unwrap().clone(),
            matches!(object.properties.get("done"), Some(Value::Bool(true))))
    }

    fn check() {
        for size in [0, 1, 8, 1024] {
            let old = generator((0..size).map(Value::I32).collect());
            let new = generator((0..size).map(Value::I32).collect());
            for _ in 0..size + 2 {
                let (old_value, old_done) = result(original_next(&old));
                let (new_value, new_done) = result(gen_next(&new));
                assert_eq!(old_done, new_done);
                match (old_value, new_value) {
                    (Value::I32(old), Value::I32(new)) => assert_eq!(old, new),
                    (Value::Undefined, Value::Undefined) => {}
                    _ => panic!("element mismatch"),
                }
            }
        }
        let item = Arc::new(Mutex::new(Object::new()));
        let source = generator(vec![Value::Object(item.clone()), Value::Undefined]);
        let (Value::Object(actual), false) = result(gen_next(&source)) else { panic!("identity lost") };
        assert!(Arc::ptr_eq(&item, &actual));
        assert!(matches!(result(gen_next(&source)), (Value::Undefined, false)));
        assert!(matches!(result(gen_next(&source)), (Value::Undefined, true)));
        assert!(source.try_lock().is_ok());
        let bare = generator(vec![Value::Object(item.clone()), Value::Undefined, Value::I32(42)]);
        let (Value::Object(actual), false) = gen_step(&bare) else { panic!("bare-step identity lost") };
        assert!(Arc::ptr_eq(&item, &actual));
        assert!(matches!(gen_step(&bare), (Value::Undefined, false)));
        assert!(matches!(gen_step(&bare), (Value::I32(42), false)));
        assert!(matches!(gen_step(&bare), (Value::Undefined, true)));
        assert!(matches!(gen_step(&bare), (Value::Undefined, true)));
        assert!(bare.try_lock().is_ok());
        let source = generator(vec![Value::I32(1)]);
        source.lock().unwrap().properties.insert(GEN_POS.into(), Value::I32(-1));
        assert!(matches!(result(gen_next(&source)), (Value::Undefined, true)));
        println!("generator order, exhaustion, undefined elements, identity, negative cursor and lock release passed");
    }

    pub fn run() {
        use std::{hint::black_box, time::Instant};
        check();
        for size in [8, 1024, 2048] {
            let old = (0..3).map(|_| generator((0..size).map(Value::I32).collect())).collect::<Vec<_>>();
            let new = (0..3).map(|_| generator((0..size).map(Value::I32).collect())).collect::<Vec<_>>();
            let start = Instant::now();
            for source in &old {
                for _ in 0..size + 1 { black_box(original_next(source)); }
            }
            let original = start.elapsed();
            let start = Instant::now();
            for source in &new {
                for _ in 0..size + 1 { black_box(gen_next(source)); }
            }
            println!("3 drains of {size} items: full-snapshot={original:?}, selected-element={:?}", start.elapsed());
            let wrapped = (0..3).map(|_| generator((0..size).map(Value::I32).collect())).collect::<Vec<_>>();
            let bare = (0..3).map(|_| generator((0..size).map(Value::I32).collect())).collect::<Vec<_>>();
            let start = Instant::now();
            for source in &wrapped {
                for _ in 0..size + 1 { black_box(result(gen_next(source))); }
            }
            let wrapped = start.elapsed();
            let start = Instant::now();
            for source in &bare {
                for _ in 0..size + 1 { black_box(gen_step(source)); }
            }
            println!("3 internal drains of {size} items: wrapped/unpacked={wrapped:?}, bare-step={:?}", start.elapsed());
        }
    }
}

fn main() { implementation::run(); }
