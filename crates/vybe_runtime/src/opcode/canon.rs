//! Component Model canon prefix (0xF0) — RETIRED, holds ZERO opcodes.
//!
//! The CM spec (proposals/component-model/design/mvp/Binary.md §Canon
//! Definitions) defines canon built-ins as `(core func)` DEFINITIONS a
//! component wires into a core instance's imports — functions, NOT
//! instructions. The VM matches that shape: every canon built-in is a
//! VM-implemented import under module "canon" (`ImportTarget::Canon` →
//! `exec_canon_builtin`), reached via spec `call` (0x00 0x10) — exactly
//! like the jspi / wasi-threads VM-implemented imports.
//!
//! The 0xF0 instruction prefix that used to carry them is retired with
//! zero rows, so any stale bytecode carrying a 0xF0 opcode fails decode
//! loudly instead of silently aliasing. Import-name mapping lives in
//! `CanonBuiltin::by_name` (vm.rs); the old sub-values are recorded here
//! only as the retirement ledger:
//!
//! - 0x00 `lift`, 0x01 `lower` — U16 typeidx immediate became a stack arg
//!   above the value.
//! - 0x02-0x04 `resource.new/drop/rep` — were declared but NEVER
//!   dispatched (executing them always faulted). No import name is
//!   registered, so they now fail at link time instead of run time.
//! - 0x05 `task.cancel`, 0x06 `subtask.cancel`, 0x09 `task.return`,
//!   0x0A/0x0B `context.get/set`, 0x0D `subtask.drop` — migrated.
//!   0x08 (`backpressure.set`) was already retired: dropped from
//!   Binary.md in favor of the inc/dec counter form.
//! - 0x0C `thread.yield`, 0x26 `thread.index`, 0x27 `thread.new-indirect`
//!   — declared, never dispatched; unregistered (fail at link time).
//! - 0x0E-0x14 stream family — `stream.new/read/write/cancel-read/
//!   drop-readable/drop-writable` migrated; `stream.cancel-write` (0x12)
//!   was never dispatched; unregistered.
//! - 0x15-0x1B future family — `future.new/drop-readable/drop-writable`
//!   migrated; `future.read/write/cancel-read/cancel-write` were never
//!   dispatched; unregistered.
//! - 0x1C-0x1E `error-context.*` — never dispatched; unregistered.
//! - 0x1F-0x23 waitable family — `waitable-set.new/wait/poll` and
//!   `waitable.join` migrated; `waitable-set.drop` (0x22) was never
//!   dispatched; unregistered.
//! - 0x24/0x25 `backpressure.inc/dec` — migrated.

use super::opcode_category;

opcode_category! {
    // ZERO rows: prefix 0xF0 decodes to None. See the module doc — canon
    // built-ins are imports, not instructions.
}
