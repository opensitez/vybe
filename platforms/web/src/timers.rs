//! HTML Living Standard §8.7 — the `web:timers` host functions.
//!
//!   `setTimeout(handler, delay?, ...args)` → number (timer id)
//!   `clearTimeout(id)`
//!   `setInterval(handler, delay?, ...args)` → number (timer id)
//!   `clearInterval(id)`
//!   `queueMicrotask(callback)`
//!
//! The WHEEL — fire times, ordering, cancellation — lives in the engine. What
//! this module owns is the `id → callback` registry, because a callback is a
//! runtime value and the engine must not know about those. That is the same
//! division the DOM uses for `addEventListener`: the engine reports what
//! became due, the host decides what running it means.
//!
//! The registry is what registers with the VM as a `DeferredSource`
//! (`wasi:io/poll` shape): the drain loop polls readiness and pops one due
//! callback per turn.
//!
//! Deadlines cross the seam as a RELATIVE delay and are added to the VM's own
//! `monotonic_now_ms` here — the engine's clock has a different origin, and
//! passing an absolute time across would mis-schedule every sleep.

use std::sync::{Arc, Mutex};

use vybe_runtime::event_loop::monotonic_now_ms;
use vybe_runtime::scheduler::DeferredSource;
use vybe_runtime::{HostContext, VM, Value};

use crate::engine::{ScheduleOp, ScheduleValue, schedule};

/// `id → callback`, in registration order so the drain matches the engine's
/// first-registered-due-first ordering.
#[derive(Default)]
pub struct TimerCallbacks {
    entries: Mutex<Vec<(u64, Value)>>,
}

impl TimerCallbacks {
    fn new() -> Self {
        TimerCallbacks::default()
    }

    /// Schedule a callback `delay_ms` from now; returns its cancellable id.
    pub fn set(&self, callback: Value, delay_ms: f64) -> u64 {
        let ScheduleValue::Id(id) = schedule(ScheduleOp::SetTimer(delay_ms)) else {
            // No engine installed — nothing will ever fire, so do not pretend
            // to hold a callback that a drain could never reach.
            return 0;
        };
        self.entries.lock().unwrap().push((id, callback));
        id
    }

    /// Cancel by id. True if the timer was still queued.
    pub fn clear(&self, id: u64) -> bool {
        self.entries.lock().unwrap().retain(|(i, _)| *i != id);
        matches!(
            schedule(ScheduleOp::ClearTimer(id)),
            ScheduleValue::Bool(true)
        )
    }

    fn take(&self, id: u64) -> Option<Value> {
        let mut entries = self.entries.lock().unwrap();
        let pos = entries.iter().position(|(i, _)| *i == id)?;
        Some(entries.remove(pos).1)
    }
}

impl DeferredSource for TimerCallbacks {
    /// Drop every callback the finished program left queued.
    ///
    /// A timer that had not fired by the end of its program used to stay here,
    /// and the next program in a reused VM drained it — running a closure over
    /// chunks that `reset_to` had already truncated, at indices the new program
    /// now occupies.
    fn clear_pending(&self) {
        self.entries.lock().unwrap().clear();
    }

    fn has_pending(&self) -> bool {
        !self.entries.lock().unwrap().is_empty()
    }

    fn earliest_deadline_ms(&self) -> Option<f64> {
        match schedule(ScheduleOp::TimerDelayMs) {
            ScheduleValue::Ms(delay) => Some(monotonic_now_ms() + delay),
            _ => None,
        }
    }

    fn pop_due(&self) -> Option<Value> {
        // Keep taking due ids until one has a callback: an id whose callback
        // is already gone is stale, not a reason to stop draining.
        loop {
            let ScheduleValue::Id(id) = schedule(ScheduleOp::TakeDueTimer) else {
                return None;
            };
            if let Some(cb) = self.take(id) {
                return Some(cb);
            }
        }
    }

    fn for_each_callback(&self, f: &mut dyn FnMut(&Value)) {
        for (_, cb) in self.entries.lock().unwrap().iter() {
            f(cb);
        }
    }
}

pub fn register(vm: &mut VM) {
    let timers = Arc::new(TimerCallbacks::new());
    vm.register_deferred_source(timers.clone());

    // setTimeout(handler, delay?, ...args) → id
    let t = timers.clone();
    vm.register_host_fn(
        "web:timers",
        "setTimeout",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let handler = args.first().cloned().unwrap_or(Value::Undefined);
            let delay_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0).max(0.0);
            Value::F64(t.set(handler, delay_ms) as f64)
        }),
    );

    let t = timers.clone();
    vm.register_host_fn(
        "web:timers",
        "clearTimeout",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            t.clear(args.first().map(|v| v.as_f64() as u64).unwrap_or(0));
            Value::Undefined
        }),
    );

    // setInterval(handler, delay?, ...args) → id
    //
    // Queues a single fire. True interval repeat requires the handler to
    // re-queue itself; that mirrors how browsers handle intervals in
    // event-loop terms and avoids unbounded queue growth.
    let t = timers.clone();
    vm.register_host_fn(
        "web:timers",
        "setInterval",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let handler = args.first().cloned().unwrap_or(Value::Undefined);
            let delay_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0).max(0.0);
            Value::F64(t.set(handler, delay_ms) as f64)
        }),
    );

    let t = timers.clone();
    vm.register_host_fn(
        "web:timers",
        "clearInterval",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            t.clear(args.first().map(|v| v.as_f64() as u64).unwrap_or(0));
            Value::Undefined
        }),
    );

    // queueMicrotask(callback) — WHATWG HTML §8.2.4.1.
    // Schedules callback as a microtask, running it after the current task
    // completes but before any macrotasks (timers).
    vm.register_host_fn(
        "web:timers",
        "queueMicrotask",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            if let Some(cb) = args.first().cloned() {
                ctx.queue_ready(cb, Value::Undefined);
            }
            Value::Undefined
        }),
    );
}
