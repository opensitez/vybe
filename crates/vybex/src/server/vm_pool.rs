//! Warm VM pool for `--serve` — the "Phase 2: VM pool" this module's sibling
//! promised at `script.rs:3`.
//!
//! A request handler is one of the three workloads `worker.rs` was written
//! for ("a test suite, a request handler, a serverless invocation"). Serving a
//! script used to cost a whole `VM::new()` + plugin + adapter registration
//! before a single line of the script ran — roughly 0.2s of pure setup on
//! every request, on top of compiling the file. Here that is paid once per
//! pool thread; a request costs a `reset_to` instead.
//!
//! ## Why its own threads and not `spawn_blocking`
//!
//! The obvious cheap version — hang the warm VM off a thread-local on tokio's
//! blocking pool — defeats itself. Tokio reaps idle blocking threads after
//! `thread_keep_alive` (10s by default), so a dev server poked by a human
//! every half minute boots a cold VM on *every* request and the pool buys
//! nothing in exactly the case it exists for; a traffic burst meanwhile grows
//! the pool toward `max_blocking_threads` (512), each booting and snapshotting
//! its own VM. Thread-local destructor order is unspecified on top of that, so
//! a VM torn down at thread exit can reach an already-destroyed `heap` or
//! resource-store TLS.
//!
//! So: a fixed set of long-lived threads owned by the server, fed by a
//! channel. Each owns one VM for the process's lifetime, and nothing else runs
//! on it.
//!
//! ## Isolation
//!
//! In a test runner a leak between jobs is a flaky test. Here it is one
//! request's file descriptors, session or key material visible to the next
//! request — a different class of bug — so this path is worth distrusting on
//! purpose. `--cold` restores the fresh-VM-per-request behaviour and is the
//! control to diff against.
//!
//! A panic mid-job is the sharpest version of that risk: a VM that panicked
//! *during its reset* is half-rolled-back, and serving the next request from
//! it would hand over whatever survived. Such a thread does not continue — it
//! spawns its own replacement (with a fresh VM, and fresh thread-locals) and
//! exits.
//!
//! ## Two consequences of a bounded pool, both deliberate
//!
//! The pool size is also the ceiling on concurrent script requests; the rest
//! queue. `spawn_blocking` had no such ceiling (it grows to 512), so two
//! behaviours changed and neither should be a surprise:
//!
//! - **A hung script holds its slot forever.** The 504 in `script::serve`
//!   releases the CLIENT, not the thread — its own message says so — and unlike
//!   a panic there is nothing to catch. N hung scripts wedge a pool of N. A
//!   watchdog that retires an over-deadline VM is the fix and is not this
//!   change.
//! - **`timeout_secs` starts when the request arrives, not when it is
//!   dequeued.** Under enough load a queued request can 504 without having run.
//!   Raising `--pool` is the lever; measuring queue depth separately from run
//!   time is the honest fix.

use std::path::PathBuf;
use std::sync::{Arc, Mutex, mpsc};

use vybe_platform_node::http::RequestContext;
use vybe_runtime::capabilities::Capabilities;

use super::compile_cache::CompileCache;

/// PHP compilation of large combined files (many `require_once` classes) uses
/// significant Rust stack depth. `mod.rs` gives tokio's blocking workers 32 MB
/// for exactly this reason; our own threads default to 2 MB and would overflow
/// where the old path did not.
const STACK_BYTES: usize = 32 * 1024 * 1024;

/// One unit of work: run `script` against a warm VM with `ctx` installed.
pub struct Job {
    pub script: PathBuf,
    pub ctx: Arc<RequestContext>,
}

pub struct VmPool {
    jobs: mpsc::Sender<Job>,
}

impl VmPool {
    /// Spawn `size` warm VM threads. Returns immediately; the threads boot
    /// their VMs in parallel and pick up work as they become ready.
    ///
    /// The compile cache is shared by every thread, not per-thread: a compiled
    /// chunk set is position-independent (see `compile_cache`), so which VM
    /// installs it is irrelevant, and a per-thread cache would need N misses to
    /// warm the same file N times.
    pub fn new(size: usize, caps: Capabilities, cache: Option<Arc<CompileCache>>) -> Self {
        let (tx, rx) = mpsc::channel::<Job>();
        // One queue, many consumers. A worker holds the lock only across
        // `recv()` and releases it the moment a job arrives, so the others
        // queue on the mutex rather than on the channel — the handoff is
        // serialised, the work is not.
        let rx = Arc::new(Mutex::new(rx));
        for index in 0..size {
            spawn_worker(index, Arc::clone(&rx), caps.clone(), cache.clone());
        }
        Self { jobs: tx }
    }

    /// Hand a request to the pool. Fails only if every worker thread is gone,
    /// which the caller answers with a 500 rather than hanging the client.
    pub fn submit(&self, job: Job) -> Result<(), &'static str> {
        self.jobs.send(job).map_err(|_| "vm pool is not running")
    }
}

fn spawn_worker(
    index: usize,
    rx: Arc<Mutex<mpsc::Receiver<Job>>>,
    caps: Capabilities,
    cache: Option<Arc<CompileCache>>,
) {
    let builder = std::thread::Builder::new()
        .name(format!("vybex-vm-{index}"))
        .stack_size(STACK_BYTES);
    let spawned = builder.spawn(move || worker_loop(index, rx, caps, cache));
    if let Err(e) = spawned {
        eprintln!("[vybex] FATAL: could not spawn VM worker {index}: {e}");
    }
}

fn worker_loop(
    index: usize,
    rx: Arc<Mutex<mpsc::Receiver<Job>>>,
    caps: Capabilities,
    cache: Option<Arc<CompileCache>>,
) {
    // The stdout→response binding is part of THIS embedder's baseline: bound
    // once here, it is a host fn every later reset preserves.
    let (mut vm, baseline) =
        match crate::warm::boot_with(&caps, super::script::register_response_stdout) {
            Ok(pair) => pair,
            Err(e) => {
                eprintln!("[vybex] FATAL: VM worker {index} failed to boot: {e}");
                return;
            }
        };

    loop {
        let job = {
            let guard = match rx.lock() {
                Ok(g) => g,
                // Poisoned means another worker panicked while holding the
                // lock — the queue itself is still sound, so keep serving.
                Err(poisoned) => poisoned.into_inner(),
            };
            match guard.recv() {
                Ok(job) => job,
                // Sender dropped: the server is shutting down.
                Err(_) => return,
            }
        };

        let ctx = Arc::clone(&job.ctx);
        let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            // BEFORE the job, never after: the previous request's response may
            // still have been draining when its job function returned, and the
            // reset drops the `wasi:http` tables that drain reads from.
            crate::warm::reset(&mut vm, &baseline);
            let _guard = vybe_platform_node::http::install_context(Arc::clone(&job.ctx));
            super::script::run_request(&mut vm, &job.script, &job.ctx, &caps, cache.as_ref());
        }));

        if outcome.is_err() {
            // The panic message itself has already gone to stderr via the
            // default hook. What matters here is that this VM is now of
            // unknown shape — possibly mid-reset — so it never serves again.
            eprintln!(
                "[vybex] VM worker {index} panicked serving {}; retiring this VM and booting a replacement",
                job.script.display()
            );
            // The client is still waiting on a channel nobody will write to.
            let mut response = match ctx.response.lock() {
                Ok(r) => r,
                Err(poisoned) => poisoned.into_inner(),
            };
            if !response.headers_sent {
                response.status = 500;
                response.headers.push((
                    "Content-Type".to_string(),
                    "text/plain; charset=utf-8".to_string(),
                ));
                response.write_bytes(
                    b"500 Internal Server Error\n\nThe script crashed the VM worker.\n".to_vec(),
                );
            }
            response.end();
            drop(response);

            spawn_worker(index, rx, caps, cache);
            return;
        }
    }
}
