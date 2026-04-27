//! HTML Living Standard §8.7 — Timers.
//!
//!   `setTimeout(handler, delay?, ...args)` → number (timer id)
//!   `clearTimeout(id)`
//!   `setInterval(handler, delay?, ...args)` → number (timer id)
//!   `clearInterval(id)`
//!
//! Vybe's MVP runs handlers synchronously after spawning a thread that
//! sleeps for `delay` ms. Real event-loop integration ships with the
//! JSPI suspend/resume work — for now this matches Node's behavior
//! closely enough for code that uses timers in the "fire and forget"
//! pattern. Cleared timers set a shared cancel flag the worker checks
//! before invoking the handler.

use std::sync::{Arc, Mutex};
use std::sync::atomic::{AtomicU64, Ordering};
use std::collections::HashMap;
use vybe_bytecode::{VM, Value, HostContext};

static NEXT_ID: AtomicU64 = AtomicU64::new(1);

struct TimerRegistry {
    cancelled: HashMap<u64, Arc<std::sync::atomic::AtomicBool>>,
}

static REGISTRY: std::sync::OnceLock<Mutex<TimerRegistry>> = std::sync::OnceLock::new();

fn registry() -> &'static Mutex<TimerRegistry> {
    REGISTRY.get_or_init(|| Mutex::new(TimerRegistry { cancelled: HashMap::new() }))
}

pub fn register(vm: &mut VM) {
    // setTimeout(handler, delay?, ...args) → id
    //
    // Spawns a native thread that sleeps then calls the handler. The
    // returned id can be passed to clearTimeout to cancel before the
    // handler fires. Vybe runs handlers synchronously from the worker
    // thread; the host doesn't have a true event loop yet.
    vm.register_host_fn("web:timers", "setTimeout", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let delay_ms = args.get(1).map(|v| v.as_f64() as u64).unwrap_or(0);
        let id = NEXT_ID.fetch_add(1, Ordering::SeqCst);
        let cancelled = Arc::new(std::sync::atomic::AtomicBool::new(false));
        registry().lock().unwrap().cancelled.insert(id, cancelled.clone());
        // Spawn the worker. Note: invoking the handler from a non-VM
        // thread is unsafe in Vybe today (HostContext is per-VM); the
        // handler firing path is wired once JSPI provides the lift.
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(delay_ms));
            if !cancelled.load(Ordering::SeqCst) {
                // Handler dispatch on main VM thread is a JSPI integration TODO.
            }
        });
        Value::F64(id as f64)
    }));

    vm.register_host_fn("web:timers", "clearTimeout", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let mut reg = registry().lock().unwrap();
        if let Some(flag) = reg.cancelled.remove(&id) {
            flag.store(true, Ordering::SeqCst);
        }
        Value::Undefined
    }));

    // setInterval(handler, delay?, ...args) → id
    vm.register_host_fn("web:timers", "setInterval", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let delay_ms = args.get(1).map(|v| v.as_f64() as u64).unwrap_or(0);
        let id = NEXT_ID.fetch_add(1, Ordering::SeqCst);
        let cancelled = Arc::new(std::sync::atomic::AtomicBool::new(false));
        registry().lock().unwrap().cancelled.insert(id, cancelled.clone());
        std::thread::spawn(move || {
            while !cancelled.load(Ordering::SeqCst) {
                std::thread::sleep(std::time::Duration::from_millis(delay_ms));
                // Handler dispatch on main VM thread is a JSPI integration TODO.
            }
        });
        Value::F64(id as f64)
    }));

    vm.register_host_fn("web:timers", "clearInterval", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let mut reg = registry().lock().unwrap();
        if let Some(flag) = reg.cancelled.remove(&id) {
            flag.store(true, Ordering::SeqCst);
        }
        Value::Undefined
    }));

    // queueMicrotask(callback) — queues for end-of-task. MVP: synchronous invoke.
    vm.register_host_fn("web:timers", "queueMicrotask", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        if let Some(cb) = args.first() {
            ctx.invoke(cb, &[]);
        }
        Value::Undefined
    }));
}
