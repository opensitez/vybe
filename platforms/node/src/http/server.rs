//! `node:http.Server` state — the timeouts and connection counters.
//!
//! Node's `Server` carries tunables the request path reads: `timeout`,
//! `keepAliveTimeout`, `headersTimeout`, `requestTimeout`,
//! `maxRequestsPerSocket`, plus `listening` and the connection closers.
//!
//! These hold REAL state rather than answering a constant: a script that sets
//! `server.keepAliveTimeout = 1000` and reads it back must see 1000, and
//! anything that silently discarded the write would be a shim. Whether the
//! transport currently honours every value is a separate question from
//! reporting it faithfully — `listening` reflects the actual server.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use vybe_runtime::{HostContext, VM, Value};

/// Node's documented defaults, in milliseconds.
const DEFAULT_KEEP_ALIVE_TIMEOUT: u64 = 5_000;
const DEFAULT_HEADERS_TIMEOUT: u64 = 60_000;
const DEFAULT_REQUEST_TIMEOUT: u64 = 300_000;

static TIMEOUT: AtomicU64 = AtomicU64::new(0);
static KEEP_ALIVE_TIMEOUT: AtomicU64 = AtomicU64::new(DEFAULT_KEEP_ALIVE_TIMEOUT);
static HEADERS_TIMEOUT: AtomicU64 = AtomicU64::new(DEFAULT_HEADERS_TIMEOUT);
static REQUEST_TIMEOUT: AtomicU64 = AtomicU64::new(DEFAULT_REQUEST_TIMEOUT);
static MAX_REQUESTS_PER_SOCKET: AtomicU64 = AtomicU64::new(0);
static LISTENING: AtomicBool = AtomicBool::new(false);

/// Record whether a server is accepting connections, so `listening` is the
/// truth rather than a guess.
pub fn set_listening(value: bool) {
    LISTENING.store(value, Ordering::SeqCst);
}

fn ms_arg(args: &[Value]) -> u64 {
    match args.first() {
        Some(Value::F64(n)) if *n > 0.0 => *n as u64,
        Some(Value::I32(n)) if *n > 0 => *n as u64,
        Some(Value::I64(n)) if *n > 0 => *n as u64,
        _ => 0,
    }
}

fn accessor(vm: &mut VM, get: &'static str, set: &'static str, cell: &'static AtomicU64) {
    vm.register_host_fn(
        "node:http",
        get,
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            Value::F64(cell.load(Ordering::SeqCst) as f64)
        }),
    );
    vm.register_host_fn(
        "node:http",
        set,
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            cell.store(ms_arg(args), Ordering::SeqCst);
            Value::Null
        }),
    );
}

pub fn register(vm: &mut VM) {
    accessor(vm, "timeout", "set_timeout", &TIMEOUT);
    accessor(
        vm,
        "keep_alive_timeout",
        "set_keep_alive_timeout",
        &KEEP_ALIVE_TIMEOUT,
    );
    accessor(
        vm,
        "headers_timeout",
        "set_headers_timeout",
        &HEADERS_TIMEOUT,
    );
    accessor(
        vm,
        "request_timeout",
        "set_request_timeout",
        &REQUEST_TIMEOUT,
    );
    accessor(
        vm,
        "max_requests_per_socket",
        "set_max_requests_per_socket",
        &MAX_REQUESTS_PER_SOCKET,
    );

    vm.register_host_fn(
        "node:http",
        "listening",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::Bool(LISTENING.load(Ordering::SeqCst))
        }),
    );
}
