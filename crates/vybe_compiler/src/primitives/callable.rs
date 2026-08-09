//! Shared callable invocation emitters.
//!
//! This is the common boundary for values that are *called*: lambdas,
//! function references, method references and delegate values. Object
//! interception belongs in the proxy primitive; proxy traps can reuse this
//! module because each trap is itself a callable.

use vybe_runtime::Chunk;

use super::Compiler;

/// Callable invocation flavor. The argument count excludes the callable value
/// already on the stack, matching `CALL_REF` and ordinary language syntax.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InvokeKind {
    Direct,
    MulticastDelegate,
}

/// Invoke the callable value currently below `arg_count` arguments on the
/// stack. Stack before: `[callable, arg0, ...]`; after: `[result]`.
pub fn emit_invoke(
    chunks: &mut [Chunk],
    current: usize,
    arg_count: u8,
    kind: InvokeKind,
    line: u32,
) {
    match kind {
        InvokeKind::Direct => emit_direct_invoke(chunks, current, arg_count, line),
        InvokeKind::MulticastDelegate => {
            emit_multicast_delegate_invoke(chunks, current, arg_count, line)
        }
    }
}

/// Direct lambda/function-reference call. The callable and its arguments must
/// already be on the stack.
pub fn emit_direct_invoke(chunks: &mut [Chunk], current: usize, arg_count: u8, line: u32) {
    emit_direct_invoke_chunk(&mut chunks[current], arg_count, line);
}

/// Direct lambda/function-reference call for helpers that are already building
/// a function chunk rather than writing into the active chunk vector.
pub fn emit_direct_invoke_chunk(chunk: &mut Chunk, arg_count: u8, line: u32) {
    crate::primitives::functions::emit_call(chunk, arg_count, line);
}

/// .NET-style multicast delegate call. The public callable contract keeps
/// `arg_count` excluding the delegate; the underlying delegate emitter uses an
/// older stack-width convention, so adapt it here once.
pub fn emit_multicast_delegate_invoke(
    chunks: &mut [Chunk],
    current: usize,
    arg_count: u8,
    line: u32,
) {
    crate::primitives::delegates::emit_invoke(chunks, current, arg_count.saturating_add(1), line);
}

impl Compiler {
    pub(crate) fn emit_callable_invoke(&mut self, arg_count: u8, kind: InvokeKind) {
        emit_invoke(&mut self.chunks, self.current, arg_count, kind, self.line);
    }

    pub(crate) fn emit_direct_callable_invoke(&mut self, arg_count: u8) {
        self.emit_callable_invoke(arg_count, InvokeKind::Direct);
    }

    pub(crate) fn emit_multicast_delegate_invoke(&mut self, arg_count: u8) {
        self.emit_callable_invoke(arg_count, InvokeKind::MulticastDelegate);
    }
}
