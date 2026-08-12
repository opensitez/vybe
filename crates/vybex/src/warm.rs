//! The warm-VM boot/reset pair, in ONE place.
//!
//! Two callers run many programs against a REUSED VM — `vybex --worker` (one
//! job per stdin line) and `vybex --serve` (one script per HTTP request) — and
//! both need the identical boot sequence and the identical reset. They had
//! their own copies, and the copies had already drifted: the server's boot was
//! missing `heap::enable_tracking()` and `prime_shared_prototypes()`, so a
//! served script's mutation of `Object.prototype` was outside the tracked heap
//! and survived into whoever was served next.
//!
//! That is the argument for this module. Boot and reset are not two lists of
//! calls, they are one contract with two halves — everything the boot puts in
//! must be something the reset can roll back to — and a contract with three
//! copies is a contract nobody is enforcing.
//!
//! `cli.rs`'s one-shot path deliberately does NOT use this. It runs one
//! program and exits: it has no baseline to return to, and tracking every
//! allocation to enable a rollback that never happens is pure cost.

use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::vm::VmSnapshot;

/// Boot a VM to the warm baseline and snapshot it.
///
/// Returns the VM together with the snapshot every later [`reset`] restores
/// to. The two are returned as a pair on purpose: a baseline taken at any
/// other moment than the end of this function is a different baseline, and
/// the whole isolation guarantee is "the tenant cannot tell it wasn't first".
pub fn boot(caps: &Capabilities) -> Result<(VM, VmSnapshot), String> {
    boot_with(caps, |_| {})
}

/// [`boot`] with an embedder hook that runs BEFORE the snapshot is taken.
///
/// `--serve` uses it to bind stdout to the HTTP response body. Host functions
/// survive a reset either way (see `VM::reset_to`'s contract), so registering
/// before the snapshot is not strictly required — but "everything the embedder
/// installs is in the baseline" is the rule that makes the baseline readable,
/// and an embedder that needs a real global rather than a host fn would break
/// without it.
pub fn boot_with(
    caps: &Capabilities,
    extra: impl FnOnce(&mut VM),
) -> Result<(VM, VmSnapshot), String> {
    // FIRST statement, before any allocation on this thread: boot-time objects
    // allocated outside the registry are invisible to `heap::restore`, so a
    // tenant's mutation of one of them could not be rolled back.
    vybe_runtime::heap::enable_tracking();

    let mut vm = VM::new();
    crate::cli::register_plugins(&mut vm, caps);
    crate::server::programmatic::register(&mut vm);
    crate::adapters::register_all(&mut vm)
        .map_err(|e| format!("adapter registration failed: {e}"))?;
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);

    // Force the shared prototypes into the tracked heap so a program that
    // mutates `Object.prototype` cannot leak into the next one. Lazily-built
    // prototypes would otherwise be created by the FIRST tenant to touch them,
    // i.e. after the snapshot, and `collect_since` would gut them for everyone.
    vybe_platform_ecma::prime_shared_prototypes();

    extra(&mut vm);

    let baseline = vm.snapshot();
    Ok((vm, baseline))
}

/// Roll a VM back to its warm baseline, ready for the next tenant.
///
/// The `reset_host_globals()` call is the other half of the pair: `reset_to`
/// owns everything reachable from the VM, and the wasi plugin's remaining
/// process-level surfaces (sql, sockets) are cleared by their own hook.
/// Everything else a plugin allocates now lives in the VM-owned resource store
/// and is dropped inside `reset_to` itself.
///
/// Call this BEFORE running a job, never after. Cleaning up afterwards looks
/// tidier and is wrong: the response half of a request may still be draining
/// when the job function returns, and the reset drops the `wasi:http` tables it
/// reads from.
pub fn reset(vm: &mut VM, baseline: &VmSnapshot) {
    vm.reset_to(baseline);
    vybe_platform_wasi::reset_host_globals();
}
