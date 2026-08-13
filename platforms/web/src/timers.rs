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

/// One scheduled callback.
///
/// `handle` is the id the GUEST holds and cancels with; `pending` is the
/// engine timer currently outstanding. For a one-shot the two are equal for
/// the timer's whole life. For a repeating timer they diverge after the first
/// fire, because the engine mints a fresh id per scheduled instant while
/// §8.6's `clearInterval(id)` must keep working on the id `setInterval`
/// returned — so the guest-facing handle is pinned to the first one.
struct Timer {
    handle: u64,
    pending: u64,
    callback: Value,
    /// `Some(delay)` for `setInterval`, `None` for `setTimeout`.
    repeat_ms: Option<f64>,
}

/// `id → callback`, in registration order so the drain matches the engine's
/// first-registered-due-first ordering.
#[derive(Default)]
pub struct TimerCallbacks {
    entries: Mutex<Vec<Timer>>,
}

impl TimerCallbacks {
    fn new() -> Self {
        TimerCallbacks::default()
    }

    /// Schedule a callback `delay_ms` from now; returns its cancellable id.
    ///
    /// `repeat` distinguishes §8.6's two entry points. It is the ONLY
    /// difference between them: "run the steps again, after the same timeout"
    /// is the whole of what `setInterval` adds.
    pub fn set(&self, callback: Value, delay_ms: f64, repeat: bool) -> u64 {
        let ScheduleValue::Id(id) = schedule(ScheduleOp::SetTimer(delay_ms)) else {
            // No engine installed — nothing will ever fire, so do not pretend
            // to hold a callback that a drain could never reach.
            return 0;
        };
        self.entries.lock().unwrap().push(Timer {
            handle: id,
            pending: id,
            callback,
            repeat_ms: repeat.then_some(delay_ms),
        });
        id
    }

    /// Cancel by the id the guest was handed. True if it was still queued.
    pub fn clear(&self, id: u64) -> bool {
        let pending = {
            let mut entries = self.entries.lock().unwrap();
            match entries.iter().position(|t| t.handle == id) {
                Some(pos) => entries.remove(pos).pending,
                // Not ours — still ask the engine, so a stale or foreign id
                // answers the same as it did before.
                None => id,
            }
        };
        matches!(
            schedule(ScheduleOp::ClearTimer(pending)),
            ScheduleValue::Bool(true)
        )
    }

    /// The callback whose engine timer just came due.
    ///
    /// A repeating timer RE-ARMS here — after the fire, never before — so at
    /// most one engine timer per interval is ever outstanding. That is what
    /// keeps a slow handler from growing the queue without bound, and it is
    /// also the spec's own ordering: the timer is rescheduled as part of
    /// running the task, not in parallel with it.
    fn take(&self, id: u64) -> Option<Value> {
        let mut entries = self.entries.lock().unwrap();
        let pos = entries.iter().position(|t| t.pending == id)?;
        let Some(delay) = entries[pos].repeat_ms else {
            return Some(entries.remove(pos).callback);
        };
        match schedule(ScheduleOp::SetTimer(delay)) {
            ScheduleValue::Id(next) => {
                entries[pos].pending = next;
                Some(entries[pos].callback.clone())
            }
            // The engine refused to reschedule, so this timer can never fire
            // again — drop it rather than leave an entry nothing will drain.
            _ => Some(entries.remove(pos).callback),
        }
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
        for t in self.entries.lock().unwrap().iter() {
            f(&t.callback);
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
            Value::F64(t.set(handler, delay_ms, false) as f64)
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
    // Repeats until `clearInterval`, per §8.6. The re-arm happens after each
    // fire (see `TimerCallbacks::take`), so exactly one engine timer per
    // interval is outstanding at a time and a slow handler cannot grow the
    // queue.
    let t = timers.clone();
    vm.register_host_fn(
        "web:timers",
        "setInterval",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let handler = args.first().cloned().unwrap_or(Value::Undefined);
            let delay_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0).max(0.0);
            Value::F64(t.set(handler, delay_ms, true) as f64)
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
