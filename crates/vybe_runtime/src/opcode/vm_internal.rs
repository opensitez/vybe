//! VM-internal opcodes (prefix 0xFF).
//! These are NOT part of the WASM specification.
//! They exist for VM performance and are lowered to standard WASM in .wasm binary output.

use super::Op;
use super::opcode_category;

impl Op {
    // 0x00 RETIRED (`const`, the custom constant-pool load). All constants
    // now emit spec instructions: i32.const/i64.const (signed LEB),
    // f32.const/f64.const (raw LE bytes), strings via wasm:string-constants
    // globals. Slot left vacant so stale bytecode fails decode loudly.
    // 0x04 RETIRED (`call_import`). Host imports are called with spec `call`
    // (0x00 0x10): u16 chunk-scoped import index + VM-internal u8 argc; the
    // .wasm writer serializes `0x10` + LEB(funcidx). Slot vacant — stale
    // bytecode fails decode loudly.
    // Branch variants
    // Immediate values
    // Type checks
    // 0xFF 0x15–0x1E: retired (DYN_* opcodes removed; replaced by wasm:js-* emitter sequences)
    // Exception handling: the custom TRY_END (0xFF 0x20) is retired. A try_table
    // is a structural block closed by spec `end` (Op::END), whose `is_try` label
    // pop removes the exception-handler group. ID 0x20 is left vacant (not reused)
    // so any legacy bytecode carrying it fails decode loudly.
    // Timers & spread
    // VM control
    // 0x23 RETIRED (`halt`). A trailing halt is the dispatch loop's
    // end-of-code path (the top frame returns top-of-stack); an early exit
    // is spec `return`; process exits with a code go through `wasi:cli/exit`.
    // The .wasm writer always lowered halt to `0x0F` (return) anyway. Slot
    // vacant — stale bytecode fails decode loudly. With this, prefix 0xFF
    // holds ZERO opcodes: the VM's instruction set is 100% WASM-spec
    // encodings.
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
    // 0x50 RETIRED (`set_type_id`). It existed to stamp an object's rtt AFTER
    // allocation, which WASM GC has no instruction for — and it needed one only
    // because the BASE constructor allocated and handed the object up, leaving
    // every derived instance carrying its parent's type. The constructor
    // protocol is inverted now: the most-derived class allocates via
    // `struct.new_default $T` and passes the receiver down, so the rtt is
    // correct from the moment the object exists. See wasmregistryfix.md §0.
    // 0x51 RETIRED (`null_none`). A GC-heap null is now the core `ref.null`
    // (0x00 0xD0) carrying `HT_NONE` as its spec heaptype immediate — this
    // opcode only existed because `ref.null` had been declared with no
    // immediate at all, leaving no way to say which heap the null belonged to.
    // Left vacant so stale bytecode fails decode loudly.
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
}

opcode_category! {
    // ZERO rows remain: every 0xFF custom opcode is retired, so every 0xFF
    // decode returns None and stale bytecode fails loudly. The invocation is
    // kept so the retirement ledger below stays with the machinery it retires.
    // 0x00 RETIRED (`const`) — spec const instructions replaced it.
    // 0x04 RETIRED (`call_import`) — spec `call` (0x00 0x10) replaced it.
    // 0x87 RETIRED (`gen.next`) — lowered to spec `resume` + `(on yield)`.
    // Row deleted so `name()` returns None and stale bytecode fails decode
    // loudly, as the retirement comments above promise.
    // 0x89–0x9C: moved to canon.rs (prefix 0xF0) with real spec binary values.
    // 0x15–0x1E: RETIRED — DYN_* opcodes removed. Slots vacant so legacy
    // bytecode carrying them fails loudly instead of silently aliasing.
    // 0x20 RETIRED (`try_end`) — a try_table closes with structural spec
    // `end`. Row deleted; vacant.
    // VM control
    // 0x23 RETIRED (`halt`) — end-of-code path / spec `return` / wasi:cli
    // exit replaced its three jobs.
    // String builtins — all removed as dead custom opcodes (string ops emit
    // wasm:js-string import calls directly).
    // 0x42–0x4A: RETIRED — ARRAY_* opcodes removed (Phase E). All callers
    // now route through ecma:array.* imports. Slots vacant; legacy bytecode
    // carrying them fails loudly rather than silently aliasing.
    // Stack-switching spec opcodes moved to core prefix 0xE0..=0xE6
    // (core_ops.rs). 0x4B..=0x4E retired/undefined.
    // 0x4F RETIRED (`promise.suspend`) — lowered to spec `suspend` with
    // AWAIT_SUSPEND_TAG. Row deleted; vacant.
    // 0x50 RETIRED (`set_type_id`) — most-derived-allocates makes the rtt
    // correct at `struct.new_default`. Row deleted; vacant.
    // 0x51 RETIRED (`null_none`) — folded into core `ref.null`'s heaptype
    // immediate. Deliberately vacant.
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
