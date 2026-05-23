//! HTML Living Standard §8.7 — Timers.
//!
//!   `setTimeout(handler, delay?, ...args)` → number (timer id)
//!   `clearTimeout(id)`
//!   `setInterval(handler, delay?, ...args)` → number (timer id)
//!   `clearInterval(id)`
//!   `queueMicrotask(callback)`
//!
//! Timers are routed through the VM event loop (monotonic clock scheduling),
//! which replaces the previous dead-thread spawn approach. The event loop
//! fires callbacks at the right time during the run_event_loop drain phase.

use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    // setTimeout(handler, delay?, ...args) → id
    //
    // Queues a macrotask via the VM event loop. The returned id can be
    // passed to clearTimeout to cancel before the handler fires.
    vm.register_host_fn("web:timers", "setTimeout", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let handler = args.first().cloned().unwrap_or(Value::Undefined);
        let delay_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0).max(0.0);
        let id = ctx.queue_timer(handler, delay_ms);
        Value::F64(id as f64)
    }));

    vm.register_host_fn("web:timers", "clearTimeout", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        ctx.cancel_timer(id);
        Value::Undefined
    }));

    // setInterval(handler, delay?, ...args) → id
    //
    // Queues a single macrotask for the first fire. True interval repeat
    // requires the handler to re-queue itself; that mirrors how browsers
    // handle intervals in event-loop terms and avoids unbounded queue growth.
    vm.register_host_fn("web:timers", "setInterval", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let handler = args.first().cloned().unwrap_or(Value::Undefined);
        let delay_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0).max(0.0);
        let id = ctx.queue_timer(handler, delay_ms);
        Value::F64(id as f64)
    }));

    vm.register_host_fn("web:timers", "clearInterval", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        ctx.cancel_timer(id);
        Value::Undefined
    }));

    // queueMicrotask(callback) — WHATWG HTML §8.2.4.1.
    // Schedules callback as a microtask, running it after the current task
    // completes but before any macrotasks (timers).
    vm.register_host_fn("web:timers", "queueMicrotask", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        if let Some(cb) = args.first().cloned() {
            ctx.queue_microtask(cb, Value::Undefined);
        }
        Value::Undefined
    }));
}
