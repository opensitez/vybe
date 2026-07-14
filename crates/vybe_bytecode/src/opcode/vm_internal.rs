//! VM-internal opcodes (prefix 0xFF).
//! These are NOT part of the WASM specification.
//! They exist for VM performance and are lowered to standard WASM in .wasm binary output.

use super::Op;
use super::opcode_category;

impl Op {
    // Constants & stack
    pub const CONST: Op = Op::new(0xFF, 0x00);
    pub const CALL_IMPORT: Op = Op::new(0xFF, 0x04);
    // Branch variants
    // Immediate values
    // Type checks
    // 0xFF 0x15–0x1E: retired (DYN_* opcodes removed; replaced by wasm:js-* emitter sequences)
    // Exception handling
    pub const TRY_END: Op = Op::new(0xFF, 0x20);
    // Timers & spread
    // VM control
    pub const HALT: Op = Op::new(0xFF, 0x23);
    // String builtins — all removed as dead custom opcodes (0x24-0x41 free).
    // String ops emit `wasm:js-string` import calls directly.
    // Array builtins (host imports in .wasm output)
    // REMOVED (Phase E): the 9 `0xFF` ARRAY_* opcodes for dynamic array
    // mutation (push, pop, slice, join, reverse, contains, indexOf,
    // concat, shift — used to live at 0xFF 0x42–0x4A). They were Vybe-
    // specific and NOT in any WASM proposal. All callers now go through
    // `ecma:array.*` imports via `common::collections::*`. The opcode
    // IDs 0x42–0x4A are left vacant (not reused) so any legacy bytecode
    // that still carries them fails decode loudly rather than silently
    // aliasing to something else.
    // Stack-switching spec opcodes (cont.new/suspend/resume/switch/
    // cont.bind/resume_throw/resume_throw_ref) live at their real
    // core-prefix spec bytes 0xE0..=0xE6 in core_ops.rs — NOT here.
    // 0x4B..=0x4F are retired and intentionally left undefined so any
    // stale bytecode carrying them fails decode loudly. 0x4F was the old
    // `PROMISE_SUSPEND` (JS `await`), now lowered to spec `suspend` with
    // `AWAIT_SUSPEND_TAG`.
    // GC extensions
    pub const SET_TYPE_ID: Op = Op::new(0xFF, 0x50);
    // Weak references
    // Multi-memory
    // Extended const
    // Typed continuations
    // String references — removed as dead custom opcodes (0x5B-0x5D free).
    // Shared GC objects
    // Component model (CANON_LIFT/CANON_LOWER moved to canon.rs with real spec values)
    // Memory64
    // JS primitive creation / testing (js-primitive-builtins proposal)
    // Narrow numeric type tests + unsigned coercions + string formatting
    // (js-primitive-builtins wiring). These give compilers direct access
    // to the declared `wasm:js-*` imports for efficient interop.
    // STR_CAST / STR_FROM_{I32,U32,I64,U64,F64} removed as dead custom
    // opcodes (0x7B-0x80 free); numeric→string now emits wasm:js-string calls.
    // ── Reference-types: typed `ref.null` variants ──────────────────
    // The core `NULL` op emits `ref.null extern` (0xD0 0x6F). These
    // emit `ref.null func` / `ref.null any` / `ref.null none` so the
    // null can be stored in slots typed anything other than externref.
    // Runtime semantics are identical — every variant pushes
    // `Value::Null`. The distinction is purely in the WASM binary.
    // `cont.bind` (0xE1) and `resume_throw` (0xE4) are real spec
    // opcodes defined at their core-prefix bytes in core_ops.rs.
    // 0x85/0x86 are retired and left undefined.
    // 0x87 is retired: the old `GEN_NEXT` (generator iterator-advance) is now
    // lowered to spec `resume` + an `(on yield)` handler. Left undefined so any
    // stale bytecode carrying it fails decode loudly.
    // All CM3 canon built-ins (stream/future/task/waitable/backpressure/context)
    // moved to canon.rs with real spec binary values on prefix 0xF0.
    // STREAM_CANCEL maps to canon stream.cancel-read (0xF0 0x11).
    // FUTURE_AWAIT has no direct canon equivalent — it's VM-level future resolution.
}

opcode_category! {
    // Constants & stack
    [0x00] r#const => U16, "const";
    [0x04] call_import => U16_U8, "call_import";
    // Branch variants
    [0x87] gen_next => None, "gen.next";
    // CM3 / WASI 0.3 async — FUTURE_AWAIT only (no direct canon equivalent)
    // 0x89–0x9C: moved to canon.rs (prefix 0xF0) with real spec binary values.
    // Immediate values
    // Type checks
    // 0x15–0x1E: RETIRED — DYN_* opcodes removed. Slots vacant so legacy
    // bytecode carrying them fails loudly instead of silently aliasing.
    // Exception handling
    [0x20] try_end => None, "try_end";
    // Timers & spread
    // VM control
    [0x23] halt => None, "halt";
    // String builtins — all removed as dead custom opcodes (string ops emit
    // wasm:js-string import calls directly).
    // 0x42–0x4A: RETIRED — ARRAY_* opcodes removed (Phase E). All callers
    // now route through ecma:array.* imports. Slots vacant; legacy bytecode
    // carrying them fails loudly rather than silently aliasing.
    // Stack-switching spec opcodes moved to core prefix 0xE0..=0xE6
    // (core_ops.rs). 0x4B..=0x4E retired/undefined.
    // JSPI
    [0x4F] promise_suspend => None, "promise.suspend";
    // GC extensions
    [0x50] set_type_id => None, "set_type_id";
    // Weak references
    // Multi-memory
    // Extended const
    // Typed continuations
    // String references
    // Shared GC objects
    // Component model (canon_lift/canon_lower moved to canon.rs)
    // Memory64
    // JS primitive creation / testing
    // Narrow numeric tests + unsigned coercions + string formatting
}
