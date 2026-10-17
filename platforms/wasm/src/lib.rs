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
//! Depends only on `vybe_runtime`'s data model (`Chunk`/`Op`/`Value`);
//! the VM never depends on this crate. Future platforms (java: .class,
//! dotnet: assemblies) follow the same shape and register their loaders
//! the same way.

pub mod disassembler;
pub(crate) mod binary_cursor;
pub(crate) mod binary_instructions;
pub mod binary_sections;
pub(crate) mod component_format;
pub mod encoding;
pub mod reader;
pub mod writer;

pub use disassembler::{wat_from_binary, write_wat};
pub use reader::{Phase, WasmError, read_wasm, read_wasm_classified};
pub use reader::component_binary::read_component_binary;
pub use writer::{WasmWriteOptions, write_wasm, write_wasm_with_options};
pub use writer::component_binary::write_component_binary;

/// Emit the declared Wasm layer. A component that is not yet representable
/// cannot silently become a core module.
pub fn write_artifact(
    chunks: &[vybe_runtime::Chunk],
    component_source: bool,
    component_binary: Option<&vybe_ast::canon::ComponentBinary>,
    options: WasmWriteOptions,
) -> Result<Vec<u8>, String> {
    if component_source {
        if options.include_vybe_metadata || options.include_name_section {
            return Err("component binary emission does not support Vybe metadata yet".into());
        }
        let binary = component_binary
            .ok_or("component binary emission does not support all source definitions yet")?;
        write_component_binary(binary)
    } else {
        if component_binary.is_some() {
            return Err("component binary metadata is missing its source marker".into());
        }
        Ok(write_wasm_with_options(chunks, options))
    }
}

/// Register this platform's binary loaders with the VM's module
/// resolver, so `ModuleResolver` can load `.wasm` files without
/// depending on this crate. Call once at host startup.
pub fn register() {
    vybe_runtime::registry::register_platform(vybe_runtime::registry::PlatformDef {
        name: "wasm",
        emit_dispatch: None,
        register_tree: None,
        namespace_constants: None,
        component_descriptor: None,
        is_descriptor_class: None,
        numeric_format_helper: None,
        read_binary_module: Some(crate::reader::read_wasm),
    });
    vybe_runtime::register_binary_loader("wasm", read_wasm);
}

/// This crate as a `vybe_runtime::Plugin` — the same single plugin type
/// languages and every other platform use. It has no VM-side surface, so
/// `init` only contributes its binary-module reader to the registry.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "wasm"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
