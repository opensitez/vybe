//! PHP dynamic type resolver adapter.
//!
//! When a class constructor global is `undefined` at runtime, PHP invokes
//! the registered `spl_autoload_register` callback (stored in the
//! `__php_autoload_callback` / `__php_autoload_callback_receiver` globals)
//! with the class name, then re-reads the constructor global. These
//! adapters emit that fallback sequence straight into the chunk.
//!
//! Mirrors the other `languages/php/emitter` adapters: chunk-based, core
//! ops only. The shared compiler routes here through the language hook and the
//! shared `dynamic_symbols` recipe owns the bytecode shape.

use vybe_bytecode::Chunk;

fn php_class_spelling(name: &str) -> String {
    name.replace('.', "\\")
}

fn php_resolver() -> vybe_compiler::compiler::dynamic_symbols::RegisteredResolver<'static> {
    vybe_compiler::compiler::dynamic_symbols::RegisteredResolver {
        callback_global: "__php_autoload_callback",
        receiver_global: "__php_autoload_callback_receiver",
    }
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
    vybe_compiler::compiler::dynamic_symbols::emit_registered_global_ref(
        chunk,
        ctor_global,
        &spelling,
        php_resolver(),
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
    vybe_compiler::compiler::dynamic_symbols::emit_registered_dynamic_global_ref(
        chunk,
        primary_ctor_global,
        fallback_ctor_global,
        &spelling,
        php_resolver(),
        line,
    );
}
