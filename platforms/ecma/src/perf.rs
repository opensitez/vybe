//! Opt-in ECMA import profiling, selected at registration time.
//!
//! Set `VYBE_ECMA_PERF=1` before registering ECMA. Disabled registrations
//! retain their original callbacks without a per-call profiling branch.
//! Counters are thread-local and include callback/guest execution inside an
//! import. Exclusive time subtracts only nested instrumented imports, not
//! all guest execution, and must not be interpreted as native ECMA CPU time.
use std::cell::RefCell;
use std::collections::HashMap;
use std::fmt::Write;
use std::sync::{Arc, OnceLock};
use std::time::{Duration, Instant};
use vybe_runtime::{HostContext, VM, Value};

type Callback = Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>;

static ENABLED: OnceLock<bool> = OnceLock::new();

#[derive(Clone, Debug, Default)]
pub struct ImportSample {
    pub name: Arc<str>,
    pub calls: u64,
    pub inclusive: Duration,
    pub exclusive: Duration,
}

struct Frame {
    start: Instant,
    children: Duration,
}

#[derive(Default)]
struct Counters {
    imports: HashMap<Arc<str>, ImportSample>,
    frames: Vec<Frame>,
}

thread_local! {
    static COUNTERS: RefCell<Counters> = RefCell::new(Counters::default());
}

/// The environment gate is read once, before the first wrapped registration.
pub fn enabled() -> bool {
    *ENABLED.get_or_init(|| matches!(std::env::var("VYBE_ECMA_PERF").as_deref(), Ok("1")))
}

struct Measurement<'a> {
    name: &'a Arc<str>,
}

impl<'a> Measurement<'a> {
    fn begin(name: &'a Arc<str>) -> Self {
        COUNTERS.with(|cell| {
            cell.borrow_mut().frames.push(Frame {
                start: Instant::now(),
                children: Duration::ZERO,
            })
        });
        Self { name }
    }
}

impl Drop for Measurement<'_> {
    fn drop(&mut self) {
        let _ = COUNTERS.try_with(|cell| {
            let mut counters = cell.borrow_mut();
            let Some(frame) = counters.frames.pop() else {
                return;
            };
            let inclusive = frame.start.elapsed();
            let exclusive = inclusive.saturating_sub(frame.children);
            if let Some(parent) = counters.frames.last_mut() {
                parent.children = parent.children.saturating_add(inclusive);
            }
            if let Some(sample) = counters.imports.get_mut(self.name.as_ref()) {
                sample.calls = sample.calls.saturating_add(1);
                sample.inclusive = sample.inclusive.saturating_add(inclusive);
                sample.exclusive = sample.exclusive.saturating_add(exclusive);
            } else {
                counters.imports.insert(
                    self.name.clone(),
                    ImportSample {
                        name: self.name.clone(),
                        calls: 1,
                        inclusive,
                        exclusive,
                    },
                );
            }
        });
    }
}

/// Wrap without changing signatures, receiver ABI, results, or throw channels.
pub(crate) fn wrap(module: &str, name: &str, callback: Callback) -> Callback {
    if !enabled() {
        return callback;
    }
    let label: Arc<str> = format!("{module}:{name}").into();
    Box::new(move |ctx, args| {
        let _measurement = Measurement::begin(&label);
        // No counter borrow or object lock survives across the callback.
        callback(ctx, args)
    })
}

pub(crate) fn register_host_fn(vm: &mut VM, module: &str, name: &str, callback: Callback) {
    vm.register_host_fn(module, name, wrap(module, name, callback));
}

pub(crate) fn register_free_fn(vm: &mut VM, module: &str, name: &str, callback: Callback) {
    // Retain the free-function declaration: profiling must not introduce a
    // receiver parameter merely because this uses the same wrapper helper.
    vm.register_free_fn(module, name, wrap(module, name, callback));
}

/// Completed import calls on the calling thread, ordered by exclusive time.
pub fn snapshot() -> Vec<ImportSample> {
    COUNTERS.with(|cell| {
        let mut samples: Vec<_> = cell.borrow().imports.values().cloned().collect();
        samples.sort_unstable_by(|a, b| {
            b.exclusive
                .cmp(&a.exclusive)
                .then_with(|| a.name.cmp(&b.name))
        });
        samples
    })
}

/// Reset between workloads, refusing to discard an active nested call stack.
pub fn reset() -> bool {
    COUNTERS.with(|cell| {
        let mut counters = cell.borrow_mut();
        if !counters.frames.is_empty() {
            return false;
        }
        counters.imports.clear();
        true
    })
}

/// Explicit reporting API for embedding callers; no implicit output in guests.
pub fn report() -> String {
    let mut output = String::from(
        "ECMA import profile (current thread; profiling overhead included)\n\
         Exclusive subtracts nested instrumented imports, not uninstrumented guest execution.\n\
         Scope: Object, Function, Reflect host/free registrations and typed Object/Array registrations.\n\
         calls  inclusive_ms  exclusive_ms  import\n",
    );
    for sample in snapshot() {
        let _ = writeln!(
            output,
            "{}  {:.3}  {:.3}  {}",
            sample.calls,
            sample.inclusive.as_secs_f64() * 1000.0,
            sample.exclusive.as_secs_f64() * 1000.0,
            sample.name
        );
    }
    output
}
