//! HTML Living Standard §8.9 — the `web:animation` host functions.
//!
//! This is the web platform's answer to "present a frame": a page never
//! pushes pixels, it asks to be called back before the next paint and draws
//! then. Everything that wanted a `present`/`swapBuffers` maps onto this —
//! an SDL game loop becomes rAF callbacks exactly the way Emscripten turns
//! `while (running)` into `emscripten_set_main_loop`.
//!
//! The frame CLOCK — cadence, ordering, cancellation — lives in the engine,
//! because a toolkit needs one to animate whether or not a runtime is
//! present. What this module owns is the `id → callback` registry, since a
//! callback is a runtime value.
//!
//! Callbacks are a [`DeferredSource`], so they drain through the SAME loop as
//! timers and jobs rather than a private event loop — the scheduling contract
//! stays declared, not invented. Per spec each callback receives a
//! `DOMHighResTimeStamp` and fires AT MOST once per registration; a callback
//! wanting the next frame re-registers, which is what makes the browser's
//! "stop drawing when nobody asks" behaviour fall out for free.

use std::sync::{Arc, Mutex};

use vybe_runtime::event_loop::monotonic_now_ms;
use vybe_runtime::scheduler::DeferredSource;
use vybe_runtime::{HostContext, VM, Value};

use crate::engine::{ScheduleOp, ScheduleValue, schedule};

/// `id → callback`, in registration order.
#[derive(Default)]
pub struct FrameCallbacks {
    entries: Mutex<Vec<(u64, Value)>>,
}

impl FrameCallbacks {
    fn new() -> Self {
        FrameCallbacks::default()
    }

    /// Register a callback for the next frame; returns its cancellable id.
    pub fn request(&self, callback: Value) -> u64 {
        let ScheduleValue::Id(id) = schedule(ScheduleOp::RequestFrame) else {
            return 0;
        };
        self.entries.lock().unwrap().push((id, callback));
        id
    }

    /// `cancelAnimationFrame`. True if the callback was still queued.
    pub fn cancel(&self, id: u64) -> bool {
        self.entries.lock().unwrap().retain(|(i, _)| *i != id);
        matches!(
            schedule(ScheduleOp::CancelFrame(id)),
            ScheduleValue::Bool(true)
        )
    }

    pub fn pending_count(&self) -> usize {
        self.entries.lock().unwrap().len()
    }

    fn take(&self, id: u64) -> Option<Value> {
        let mut entries = self.entries.lock().unwrap();
        let pos = entries.iter().position(|(i, _)| *i == id)?;
        Some(entries.remove(pos).1)
    }
}

impl DeferredSource for FrameCallbacks {
    /// Drop every frame callback the finished program left queued — same
    /// reasoning as [`crate::timers::TimerCallbacks::clear_pending`]: a
    /// `requestAnimationFrame` callback belongs to the program that requested
    /// it, and the registry outlives that program.
    fn clear_pending(&self) {
        self.entries.lock().unwrap().clear();
    }

    fn has_pending(&self) -> bool {
        !self.entries.lock().unwrap().is_empty()
    }

    fn earliest_deadline_ms(&self) -> Option<f64> {
        match schedule(ScheduleOp::FrameDelayMs) {
            ScheduleValue::Ms(delay) => Some(monotonic_now_ms() + delay),
            _ => None,
        }
    }

    fn pop_due(&self) -> Option<Value> {
        loop {
            let ScheduleValue::Id(id) = schedule(ScheduleOp::TakeDueFrame) else {
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

/// The process-wide frame-callback registry, reachable without a VM handle so
/// a window backend can ask whether anyone is animating.
pub fn callbacks() -> Arc<FrameCallbacks> {
    static CALLBACKS: std::sync::OnceLock<Arc<FrameCallbacks>> = std::sync::OnceLock::new();
    CALLBACKS
        .get_or_init(|| Arc::new(FrameCallbacks::new()))
        .clone()
}

pub fn register(vm: &mut VM) {
    let frames = callbacks();
    vm.register_deferred_source(frames.clone());

    // requestAnimationFrame(callback) → id
    {
        let f = frames.clone();
        vm.register_host_fn(
            "web:animation",
            "requestAnimationFrame",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let cb = args.first().cloned().unwrap_or(Value::Undefined);
                Value::F64(f.request(cb) as f64)
            }),
        );
    }

    // cancelAnimationFrame(id)
    {
        let f = frames.clone();
        vm.register_host_fn(
            "web:animation",
            "cancelAnimationFrame",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
                Value::Bool(f.cancel(id))
            }),
        );
    }

    // `performance.now()` — the timestamp rAF callbacks are handed, exposed
    // on its own because guests time frames with it. Answered on the VM's
    // clock, which is what every other guest-visible time uses.
    vm.register_host_fn(
        "web:animation",
        "now",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| Value::F64(monotonic_now_ms())),
    );
}
