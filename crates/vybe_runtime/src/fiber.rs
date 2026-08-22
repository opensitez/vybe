//! Fiber — a suspendable execution context.
//! When `await` is hit, the current fiber is suspended (stack + frames saved).
//! When the awaited Promise resolves, the fiber is resumed.

use crate::value::{Upvalue, Value};
use std::sync::{Arc, Mutex};

/// A suspended execution context — everything needed to resume.
#[derive(Debug)]
pub struct Fiber {
    /// Saved operand stack.
    pub stack: Vec<Value>,
    /// Saved call frames.
    pub frames: Vec<SavedFrame>,
    /// Saved open upvalues.
    pub open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    /// Saved structured-control-flow label entries. Generators that
    /// yield inside a `while` / `for` / `block` must preserve the
    /// surrounding block/loop labels so `br` / `br_if`
    /// target the right depth after resumption.
    pub label_stack: Vec<crate::vm::LabelEntry>,
    /// Saved WASM exception-handler stack. Suspending inside a try body
    /// must preserve the active catch targets so `resume_throw` and
    /// rejected JSPI resumes enter the restored fiber's structured EH path.
    pub(crate) exception_handlers: Vec<crate::vm::ExceptionHandler>,
    /// Saved active-continuation stack. A fiber captured inside a
    /// running coroutine (e.g. `await` inside a `@generator`) must
    /// bring its coroutine-context with it; restoring the fiber
    /// resurrects the nested generator state along with everything
    /// else.
    pub active_continuations: Vec<crate::vm::ActiveContinuation>,
    /// The value to push onto the stack when resuming (the await result).
    pub resume_value: Option<Value>,
    /// If set, resume by throwing this value instead of pushing resume_value.
    /// Used when a rejected promise resumes a suspended fiber — the rejection
    /// reason must be thrown (not returned) so enclosing try/catch blocks fire.
    pub resume_exception: Option<Value>,
    /// Identity of this fiber. Captured from `VM::cur_fiber_id` at `save_fiber`
    /// and restored on resume so a nested `execute_until`'s `min_depth` boundary
    /// is only honoured on the fiber it was entered on (see `VM::cur_fiber_id`).
    pub(crate) fiber_id: u64,
    /// JSPI promising boundary (async function call): the pending result
    /// Promise handed to the async fn's caller when this fiber suspended at an
    /// `await`. When the fiber finally runs to completion, the VM settles this
    /// promise with the body's outcome (fulfil on return, reject on throw).
    pub(crate) result_promise: Option<Value>,
    /// Frame-depth floors of async-function calls active on THIS fiber's frame
    /// stack (see `VM::async_floors`). Fiber-local: stack switches swap it with
    /// the rest of the execution state so floors never index another fiber's
    /// frames.
    pub(crate) async_floors: Vec<usize>,
    /// A SYNCHRONOUS `canon stream.read` / `future.read` parked mid-copy.
    ///
    /// `CanonicalABI.md` §`canon stream.{read,write}`: only the `async` variant
    /// may answer `BLOCKED`; the synchronous one must SUSPEND until the copy
    /// can proceed. The instruction that would have filled the guest's buffer
    /// has already retired by the time the fiber parks, so the copy is
    /// re-performed HOST-side at resume and its packed `CopyResult` becomes the
    /// fiber's single resume value. That is what makes suspension expressible
    /// on a resume path that pushes exactly one value.
    pub(crate) pending_copy: Option<PendingCopy>,
}

/// The copy a parked fiber still owes: enough to redo it, nothing more.
///
/// `handle` is carried alongside `end_id` because the two answer different
/// questions — `end_id` finds the data, `handle` finds the handle-table entry
/// whose `CopyState` the copy has to settle.
#[derive(Debug, Clone)]
pub struct PendingCopy {
    pub handle: u32,
    pub end_id: u64,
    pub ptr: usize,
    /// Element count, not bytes — the unit `canon stream.read` copies in.
    pub n: usize,
    pub kind: PendingCopyKind,
}

#[derive(Debug, Clone)]
pub enum PendingCopyKind {
    /// `stream<u8>` — the byte path.
    StreamBytes,
    /// `stream<T>`, `T` ≠ `u8` — elements are stored at `T`'s layout.
    StreamTyped(crate::component::ValType),
    /// `future<T>` — exactly one element, or none.
    Future(crate::component::ValType),
}

#[derive(Debug, Clone)]
pub struct SavedFrame {
    pub chunk_index: usize,
    pub ip: usize,
    pub base: usize,
    pub label_base: usize,
    pub upvalues: Vec<Arc<Mutex<Upvalue>>>,
}

impl Fiber {
    pub fn new(
        stack: Vec<Value>,
        frames: Vec<SavedFrame>,
        open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    ) -> Self {
        Fiber {
            stack,
            frames,
            open_upvalues,
            label_stack: Vec::new(),
            exception_handlers: Vec::new(),
            active_continuations: Vec::new(),
            resume_value: None,
            resume_exception: None,
            fiber_id: 0,
            result_promise: None,
            async_floors: Vec::new(),
            pending_copy: None,
        }
    }
    pub fn with_labels(mut self, labels: Vec<crate::vm::LabelEntry>) -> Self {
        self.label_stack = labels;
        self
    }
    pub(crate) fn with_fiber_id(mut self, id: u64) -> Self {
        self.fiber_id = id;
        self
    }
    pub(crate) fn with_exception_handlers(
        mut self,
        handlers: Vec<crate::vm::ExceptionHandler>,
    ) -> Self {
        self.exception_handlers = handlers;
        self
    }
    pub fn with_continuations(mut self, conts: Vec<crate::vm::ActiveContinuation>) -> Self {
        self.active_continuations = conts;
        self
    }
}
