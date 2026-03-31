//! System.Threading.Tasks, ThreadPool, Interlocked, Stopwatch, BackgroundWorker

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // Task.Run — simplified: just call the function synchronously
    vm.register_host_fn("vybe:threading", "taskRun", Box::new(|_args: &[Value]| {
        // In our synchronous VM, Task.Run just executes immediately
        // The callback is args[0], but we can't call it from a host fn
        // Return a "completed task" object
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        obj.properties.insert("status".into(), Value::String(Rc::from("RanToCompletion")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Task.Delay(ms) — sleep
    vm.register_host_fn("vybe:threading", "taskDelay", Box::new(|args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        std::thread::sleep(std::time::Duration::from_millis(ms));
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Task.FromResult(value)
    vm.register_host_fn("vybe:threading", "taskFromResult", Box::new(|args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        obj.properties.insert("result".into(), val);
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Task.CompletedTask
    vm.register_host_fn("vybe:threading", "taskCompleted", Box::new(|_args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Task")));
        obj.properties.insert("iscompleted".into(), Value::Bool(true));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Stopwatch
    vm.register_host_fn("vybe:threading", "stopwatchNew", Box::new(|_args: &[Value]| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as f64;
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Stopwatch")));
        obj.properties.insert("__start".into(), Value::F64(now));
        obj.properties.insert("isrunning".into(), Value::Bool(true));
        obj.properties.insert("elapsedmilliseconds".into(), Value::F64(0.0));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    vm.register_host_fn("vybe:threading", "stopwatchElapsed", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let start = o.properties.get("__start").map(|v| v.as_f64()).unwrap_or(0.0);
            let now = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_millis() as f64;
            return Value::F64(now - start);
        }
        Value::F64(0.0)
    }));

    // System.Random
    vm.register_host_fn("vybe:threading", "randomNew", Box::new(|_args: &[Value]| {
        let seed = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos() as u64;
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Random")));
        obj.properties.insert("__seed".into(), Value::F64(seed as f64));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    vm.register_host_fn("vybe:threading", "randomNext", Box::new(|args: &[Value]| {
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

    vm.register_host_fn("vybe:threading", "randomNextDouble", Box::new(|_args: &[Value]| {
        let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default().subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
}
