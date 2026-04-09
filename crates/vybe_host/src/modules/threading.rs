//! System.Threading.Tasks, ThreadPool, Interlocked, Stopwatch, BackgroundWorker

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Thread/Task spawning is handled by WASM stack switching opcodes
    // (cont_new + resume) emitted by compiler_common::threading.
    // No host functions needed for spawn/join — it's pure WASM.

    // Task.Run — legacy host fn for old compilers that still use call_import
    vm.register_host_fn("vybe:threading", "taskRun", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let callback = args.first().cloned().unwrap_or(Value::Null);
        let result = ctx.invoke(&callback, &[]);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        obj.properties.insert("status".into(), Value::String(Arc::from("RanToCompletion")));
        obj.properties.insert("result".into(), result);
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Task.Delay(ms) — sleep
    vm.register_host_fn("vybe:threading", "taskDelay", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        std::thread::sleep(std::time::Duration::from_millis(ms));
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Task.FromResult(value)
    vm.register_host_fn("vybe:threading", "taskFromResult", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        obj.properties.insert("result".into(), val);
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Task.CompletedTask
    vm.register_host_fn("vybe:threading", "taskCompleted", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Stopwatch.Start()
    vm.register_host_fn("vybe:threading", "stopwatchStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            let running = o.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if !running {
                let now = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_millis() as f64;
                o.properties.insert("__start".into(), Value::F64(now));
                o.properties.insert("isrunning".into(), Value::Bool(true));
            }
        }
        Value::Null
    }));

    // Stopwatch.Stop()
    vm.register_host_fn("vybe:threading", "stopwatchStop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            let running = o.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if running {
                let start = o.properties.get("__start").map(|v| v.as_f64()).unwrap_or(0.0);
                let acc = o.properties.get("__accumulated").map(|v| v.as_f64()).unwrap_or(0.0);
                let now = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_millis() as f64;
                o.properties.insert("__accumulated".into(), Value::F64(acc + (now - start)));
                o.properties.insert("isrunning".into(), Value::Bool(false));
            }
        }
        Value::Null
    }));

    // Stopwatch.ElapsedMilliseconds (property getter)
    vm.register_host_fn("vybe:threading", "stopwatchElapsed", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let acc = o.properties.get("__accumulated").map(|v| v.as_f64()).unwrap_or(0.0);
            let running = o.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if running {
                let start = o.properties.get("__start").map(|v| v.as_f64()).unwrap_or(0.0);
                let now = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_millis() as f64;
                return Value::F64(acc + (now - start));
            }
            return Value::F64(acc);
        }
        Value::F64(0.0)
    }));

    // Stopwatch constructor — capture method indices for __get_ property getters
    let elapsed_idx = *vm.host_registry.get(&("vybe:threading".into(), "stopwatchElapsed".into())).unwrap();
    vm.register_host_fn("vybe:threading", "stopwatchNew", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Stopwatch")));
        obj.properties.insert("__start".into(), Value::F64(0.0));
        obj.properties.insert("__accumulated".into(), Value::F64(0.0));
        obj.properties.insert("isrunning".into(), Value::Bool(false));
        // ElapsedMilliseconds is a .NET property — register as __get_ so struct_get auto-invokes
        let mut getter = Object::new();
        getter.kind = ObjectKind::HostFunction(elapsed_idx);
        let getter_val = Value::Object(Arc::new(Mutex::new(getter)));
        obj.properties.insert("__get_elapsedmilliseconds".into(), getter_val.clone());
        obj.properties.insert("__get_elapsed".into(), getter_val);
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // System.Random
    vm.register_host_fn("vybe:threading", "randomNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let seed = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos() as u64;
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Random")));
        obj.properties.insert("__seed".into(), Value::F64(seed as f64));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:threading", "randomNext", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let max = args.get(1).map(|v| v.as_f64() as u64).unwrap_or(i32::MAX as u64);
        let min = if args.len() > 2 {
            let a = args.get(1).map(|v| v.as_f64() as u64).unwrap_or(0);
            let b = args.get(2).map(|v| v.as_f64() as u64).unwrap_or(i32::MAX as u64);
            // randomNext(min, max)
            let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default().subsec_nanos() as u64;
            return Value::F64((a + t % (b - a).max(1)) as f64);
        } else {
            0u64
        };
        let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default().subsec_nanos() as u64;
        Value::F64((min + t % max.max(1)) as f64)
    }));

    vm.register_host_fn("vybe:threading", "randomNextDouble", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default().subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
}
