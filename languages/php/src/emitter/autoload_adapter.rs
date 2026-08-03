//! PHP autoload adapter — the `spl_autoload_*` surface.
//!
//! PHP keeps an ordered queue of autoloaders. `spl_autoload_register` appends
//! (or prepends) to it, `spl_autoload_unregister` removes by identity,
//! `spl_autoload_functions` reports it, and `spl_autoload_call` runs it by
//! hand. When a class constructor global is still `undefined` at runtime, the
//! queue runs in order until one of them defines the class.
//!
//! Every PHP-specific spelling lives here — the global holding the queue, the
//! `\` class-name separator, and `__invoke` for callable objects. The shared
//! `dynamic_symbols` recipe owns the bytecode shape and knows none of them.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use vybe_compiler::primitives::dynamic_symbols::{self, ResolverStack};

/// The queue of registered autoloaders, and the protocol member that makes an
/// object callable in PHP.
fn php_autoload_stack() -> ResolverStack<'static> {
    ResolverStack {
        stack_global: "__php_autoload_stack",
        invoke_member: Some("__invoke") }
}

fn php_class_spelling(name: &str) -> String {
    name.replace('.', "\\")
}

/// Push a reference to `ctor_global`, autoloading the class first if the
/// global is still undefined. Stack on exit: `[ctor_ref]`.
pub fn emit_constructor_ref_with_autoload(
    chunk: &mut Chunk,
    ctor_global: &str,
    autoload_name: &str,
    line: u32,
) {
    let spelling = php_class_spelling(autoload_name);
    dynamic_symbols::emit_registered_global_ref(
        chunk,
        ctor_global,
        &spelling,
        php_autoload_stack(),
        line,
    );
}

/// Like [`emit_constructor_ref_with_autoload`] but resolves a primary
/// constructor global, then an optional fallback global, before
/// autoloading. Stack on exit: `[ctor_ref]`.
pub fn emit_dynamic_constructor_ref_with_autoload(
    chunk: &mut Chunk,
    primary_ctor_global: &str,
    fallback_ctor_global: Option<&str>,
    autoload_name: &str,
    line: u32,
) {
    let spelling = php_class_spelling(autoload_name);
    dynamic_symbols::emit_registered_dynamic_global_ref(
        chunk,
        primary_ctor_global,
        fallback_ctor_global,
        &spelling,
        php_autoload_stack(),
        line,
    );
}

/// `spl_autoload_register(callable $callback, bool $throw = true, bool $prepend = false)`.
///
/// `$throw` is accepted and ignored: a bad callable is reported the same way
/// either value. The args are already on the stack, so this normalizes them to
/// the `[callable, prepend]` the shared recipe expects.
pub fn emit_spl_autoload_register(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        // `spl_autoload_register()` with no callback registers PHP's own
        // default loader, which has nothing to queue here.
        0 => {
            chunk.emit_bool_const(true, line);
            return;
        }
        1 => chunk.emit_bool_const(false, line),
        2 => {
            chunk.emit_op(Op::DROP, line); // $throw
            chunk.emit_bool_const(false, line);
        }
        _ => {
            let prepend = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, prepend, line); // $prepend
            chunk.emit_op(Op::DROP, line); // $throw
            chunk.emit_op_u16(Op::LOCAL_GET, prepend, line);
        }
    }
    dynamic_symbols::emit_resolver_register(chunk, php_autoload_stack(), line);
}

/// `spl_autoload_unregister(callable $callback)` — removes by identity.
pub fn emit_spl_autoload_unregister(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_bool_const(false, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    dynamic_symbols::emit_resolver_unregister(chunk, php_autoload_stack(), line);
}

/// `spl_autoload_functions()` — the registered autoloaders, in call order.
pub fn emit_spl_autoload_functions(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    dynamic_symbols::emit_resolver_list(chunk, php_autoload_stack(), line);
}

/// `spl_autoload_call(string $class)` — run the queue by hand. Always yields
/// `null`, like PHP: the loaders' job is the side effect of defining the class.
///
/// `resolved_global` stays `None` because the class name is a runtime value
/// here, so every loader gets a turn rather than stopping at the first hit.
/// The literal-name case is specialized by the walker, which knows the global.
pub fn emit_spl_autoload_call(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    dynamic_symbols::emit_resolver_stack_invoke(chunk, php_autoload_stack(), None, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
