//! `node:perf_hooks` — Node.js performance measurement APIs.
//!
//! Reference: <https://nodejs.org/api/perf_hooks.html>.

use std::sync::{Arc, OnceLock};
use std::time::{Instant, SystemTime, UNIX_EPOCH};
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn perf_origin() -> &'static Instant {
    static ORIGIN: OnceLock<Instant> = OnceLock::new();
    ORIGIN.get_or_init(Instant::now)
}

fn now_ms() -> f64 {
    perf_origin().elapsed().as_secs_f64() * 1000.0
}

fn time_origin_ms() -> f64 {
    let since_epoch = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs_f64() * 1000.0)
        .unwrap_or(1.0);
    // Subtract how long we've been running to get when the process started.
    since_epoch - now_ms()
}

fn empty_array() -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object {
        kind: ObjectKind::Array(vec![]),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: vec![],
    }))
}

fn mark_entry(name: &str) -> Value {
    let mut o = Object::new();
    o.properties.insert("name".into(), s(name));
    o.properties.insert("entryType".into(), s("mark"));
    o.properties
        .insert("startTime".into(), Value::F64(now_ms()));
    o.properties.insert("duration".into(), Value::F64(0.0));
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn histogram_stub() -> Value {
    let mut o = Object::new();
    o.properties.insert("min".into(), Value::F64(0.0));
    o.properties.insert("max".into(), Value::F64(0.0));
    o.properties.insert("mean".into(), Value::F64(0.0));
    o.properties.insert("stddev".into(), Value::F64(0.0));
    o.properties.insert("exceeds".into(), Value::F64(0.0));
    o.properties.insert("percentile".into(), Value::Null);
    o.properties.insert("percentiles".into(), Value::Null);
    o.properties.insert("count".into(), Value::I32(0));
    Value::Object(vybe_runtime::heap::alloc(o))
}

pub fn register(vm: &mut VM) {
    let _ = perf_origin();

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceNow",
        Box::new(|_ctx, _args| Value::F64(now_ms())),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceTimeOrigin",
        Box::new(|_ctx, _args| Value::F64(time_origin_ms().max(1.0))),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceMark",
        Box::new(|_ctx, args| {
            let name = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            mark_entry(&name)
        }),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceMeasure",
        Box::new(|_ctx, args| {
            let name = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            let mut o = Object::new();
            o.properties.insert("name".into(), s(&name));
            o.properties.insert("entryType".into(), s("measure"));
            o.properties
                .insert("startTime".into(), Value::F64(now_ms()));
            o.properties.insert("duration".into(), Value::F64(0.0));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceClearMarks",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceClearMeasures",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceGetEntries",
        Box::new(|_ctx, _args| empty_array()),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceGetEntriesByName",
        Box::new(|_ctx, _args| empty_array()),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "performanceGetEntriesByType",
        Box::new(|_ctx, _args| empty_array()),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "eventLoopUtilization",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("idle".into(), Value::F64(0.0));
            o.properties.insert("active".into(), Value::F64(0.0));
            o.properties.insert("utilization".into(), Value::F64(0.0));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "monitorEventLoopDelay",
        Box::new(|_ctx, _args| histogram_stub()),
    );

    vm.register_host_fn(
        "node:perf_hooks",
        "createHistogram",
        Box::new(|_ctx, _args| histogram_stub()),
    );

    // Stub constructors — just return an empty object
    for name in [
        "PerformanceObserver",
        "PerformanceEntry",
        "PerformanceMark",
        "PerformanceMeasure",
        "PerformanceResourceTiming",
        "PerformanceNodeTiming",
    ] {
        vm.register_host_fn(
            "node:perf_hooks",
            name,
            Box::new(|_ctx, _args| Value::Object(vybe_runtime::heap::alloc(Object::new()))),
        );
    }
}

#[allow(dead_code)]
fn _force_use(_: ObjectKind) {}
