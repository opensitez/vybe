//! `System.Threading.Tasks.TaskCompletionSource(Of T)` — the producer half of a
//! Task, shared by every .NET language.
//!
//! The type was absent from the catalog entirely, so every member answered
//! "undefined is not callable" and all twenty
//! `vb_task_completion_source_set_result` cases sat at zero. It is registered
//! here as DATA — a `ClassType` that `tree_register` mounts in the shared
//! namespace tree — so the common resolver answers it from ambient roots, the
//! same route `Lazy(Of T)` takes.
//!
//! Shape: `{__type:"TaskCompletionSource", __tcs_promise, __tcs_resolve,
//! __tcs_reject, __tcs_done}`, built from `ecma:promise.withResolvers()` —
//! ECMA's own deferred, which is precisely what a TCS is. Reusing it means the
//! task this hands out is an ordinary promise, so `Await`, `.Result`,
//! `ContinueWith` and `Task.WhenAll` all keep working through the paths they
//! already have.
//!
//! `.Task` is a READ-ONLY computed property declared in
//! `tree_register::shared_emit_accessors`, not a struct field — a field named
//! `Task` on the source would be read by a case-insensitive front end as
//! `task`, and the accessor keeps the promise reachable under the declared
//! spelling for C# too.
//!
//! ## `Set*` throws where `TrySet*` answers
//!
//! .NET draws the line at the SECOND completion: `SetResult` on an already
//! completed source throws `InvalidOperationException`, while `TrySetResult`
//! returns `False` and leaves the first result standing. Both directions share
//! one emitter and branch on `__tcs_done`, because a second copy of that rule
//! is a second place for it to drift.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";
const PROMISE_KEY: &str = "__tcs_promise";
const RESOLVE_KEY: &str = "__tcs_resolve";
const REJECT_KEY: &str = "__tcs_reject";
const DONE_KEY: &str = "__tcs_done";
const TYPE_NAME: &str = "TaskCompletionSource";




fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    vybe_compiler::primitives::errors::emit_exception_new(
        chunk,
        exception_name,
        class_slots::ValueSource::ConstStr(message.to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

/// `New TaskCompletionSource(Of T)()` and its `state` / `TaskCreationOptions`
/// overloads.
///
/// The overload arguments are accepted and dropped. `TaskCreationOptions`
/// selects a CONTINUATION SCHEDULING policy — whether a continuation runs
/// inline on the completing thread or is queued — and every continuation here
/// already goes through the microtask queue, so nothing can observe the
/// difference. Dropping is honest; storing a flag nothing reads would not be.
///
/// Stack on entry: `[]` or `[arg]` ; on exit: `[tcs]`.
pub fn emit_tcs_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    let (obj_slot, deferred_slot) = {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(2);
        (base, base + 1)
    };

    call_import(chunks, current, "ecma:promise", "withResolvers", 0, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, deferred_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    lget(chunk, obj_slot, line);
    chunk.emit_string_const(TYPE_NAME, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(TYPE_KEY),
        ValueSource::Stack,
        line,
    );

    for (field, source) in [
        (PROMISE_KEY, "promise"),
        (RESOLVE_KEY, "resolve"),
        (REJECT_KEY, "reject"),
    ] {
        lget(chunk, obj_slot, line);
        lget(chunk, deferred_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(source),
            Dest::Stack,
            line,
        );
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(field),
            ValueSource::Stack,
            line,
        );
    }

    lget(chunk, obj_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(DONE_KEY),
        ValueSource::Stack,
        line,
    );

    lget(chunk, obj_slot, line);
}

/// `tcs.Task` — the consumer half.
///
/// Stack on entry: `[tcs]` ; on exit: `[task]`.
pub fn emit_tcs_task(chunks: &mut [Chunk], current: usize, line: u32) {
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot(PROMISE_KEY),
        Dest::Stack,
        line,
    );
}

/// Which settler a completion runs, and what a repeat completion does.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Settle {
    Result,
    Exception,
    Canceled,
}

impl Settle {
    fn settler(self) -> &'static str {
        match self {
            Settle::Result => RESOLVE_KEY,
            Settle::Exception | Settle::Canceled => REJECT_KEY,
        }
    }

    /// The .NET message for completing an already-completed source.
    fn already_completed(self) -> &'static str {
        "An attempt was made to transition a task to a final state when it had already completed."
    }
}

/// The shared body behind `SetResult`/`TrySetResult`,
/// `SetException`/`TrySetException` and `SetCanceled`/`TrySetCanceled`.
///
/// `try_variant` picks the .NET contract: `TrySet*` leaves a `Bool` on the
/// stack and never throws; `Set*` leaves nothing and throws
/// `InvalidOperationException` on a second completion.
///
/// Stack on entry: `[tcs]` (Canceled) or `[tcs, value]` ; on exit: `[bool]`
/// when `try_variant`, otherwise `[]`.
pub fn emit_tcs_settle(
    chunks: &mut Vec<Chunk>,
    current: usize,
    settle: Settle,
    try_variant: bool,
    line: u32,
) {
    let (obj_slot, value_slot) = {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(2);
        (base, base + 1)
    };

    {
        let chunk = &mut chunks[current];
        if settle == Settle::Canceled {
            // `SetCanceled()` carries no value — the rejection reason is the
            // cancellation itself, minted here.
            chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
            vybe_compiler::primitives::errors::emit_exception_new(
                chunk,
                "TaskCanceledException",
                class_slots::ValueSource::ConstStr("A task was canceled.".to_string()),
                line,
            );
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        } else {
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
        }

        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(DONE_KEY),
            Dest::Stack,
            line,
        );
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        if try_variant {
            chunk.emit_if_value(line);
        } else {
            chunk.emit_if(line);
        }

        // ── not yet completed: settle, and mark it ──
        lget(chunk, obj_slot, line);
        chunk.emit_bool_const(true, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(DONE_KEY),
            ValueSource::Stack,
            line,
        );

        lget(chunk, obj_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(settle.settler()),
            Dest::Stack,
            line,
        );
        lget(chunk, value_slot, line);
    }
    // The settler is an ECMA bound function; calling it is what actually
    // transitions the promise and drains its reactions.
    //
    // ⛔ `emit_invoke`'s count INCLUDES THE DELEGATE — `callable.rs` passes
    // `arg_count.saturating_add(1)`. Passing 1 for one argument makes it pop
    // the ARGUMENT as the callable: `SetResult("R")` died with
    // `string is not callable (type: R)`.
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 2, line);
    {
        let chunk = &mut chunks[current];
        // `resolve`/`reject` answer undefined — the settle is the point.
        chunk.emit_op(Op::DROP, line);

        // ⛔ CANCELLATION HAS TO BE CARRIED, NOT DERIVED. ECMA has ONE
        // `rejected` state; .NET distinguishes `Canceled` from `Faulted`. The
        // stamp is what lets `task.IsCanceled` and `task.Status` answer
        // differently for `SetCanceled` than for `SetException` — without it
        // both read `Faulted` and `IsCanceled` is False after a cancel.
        if settle == Settle::Canceled {
            lget(chunk, obj_slot, line);
            class_slots::emit_class_get(
                chunk,
                ObjSource::Stack,
                &field_slot(PROMISE_KEY),
                Dest::Stack,
                line,
            );
            chunk.emit_string_const("Canceled", line);
            class_slots::emit_class_set(
                chunk,
                ObjSource::Stack,
                &field_slot("status"),
                ValueSource::Stack,
                line,
            );
        }

        if try_variant {
            chunk.emit_bool_const(true, line);
            chunk.emit_else(line);
            chunk.emit_bool_const(false, line);
        } else {
            chunk.emit_else(line);
            emit_throw_dotnet_exception(
                chunk,
                "InvalidOperationException",
                settle.already_completed(),
                line,
            );
        }
        chunk.emit_end(line);
    }
}
