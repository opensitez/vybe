//! WASM platform — the binary codec between Vybe `Chunk`s and
//! WebAssembly modules. First platform in the standard shape:
//!
//! - `reader/` — .wasm binary → `Vec<Chunk>` (vybe custom section for
//!   round-trip, standard-section decode for foreign modules)
//! - `writer/` — `Vec<Chunk>` → .wasm binary (sections, GC types, code,
//!   per-proposal modules, `wasm:js-*` builtin import surface)
//! - `disassembler/` — `Vec<Chunk>` → WAT text for human inspection
//! - `encoding` — shared vocabulary: magic, section IDs, LEB128
//!
//! Depends only on `vybe_bytecode`'s data model (`Chunk`/`Op`/`Value`);
//! the VM never depends on this crate. Future platforms (java: .class,
//! dotnet: assemblies) follow the same shape and register their loaders
//! the same way.

pub mod disassembler;
pub mod encoding;
pub mod reader;
pub mod writer;

pub use disassembler::write_wat;
pub use reader::read_wasm;
pub use writer::write_wasm;

/// Register this platform's binary loaders with the VM's module
/// resolver, so `ModuleResolver` can load `.wasm` files without
/// depending on this crate. Call once at host startup.
pub fn register() {
    vybe_bytecode::registry::register_platform(vybe_bytecode::registry::PlatformDef {
        name: "wasm",
        emit_dispatch: None,
        register_tree: None,
        namespace_constants: None,
        component_descriptor: None,
        is_descriptor_class: None,
        numeric_format_helper: None,
        read_binary_module: Some(crate::reader::read_wasm),
    });
    vybe_bytecode::register_binary_loader("wasm", read_wasm);
}

/// This crate as a `vybe_bytecode::Plugin` — the same single plugin type
/// languages and every other platform use. It has no VM-side surface, so
/// `init` only contributes its binary-module reader to the registry.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "wasm"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_bytecode::register_plugin!(Plugin);
