//! WASM Component Model — typed interfaces for cross-language module linking.
//!
//! A Component defines:
//! - Imports: interfaces this component needs (resolved at link time)
//! - Exports: interfaces this component provides
//!
//! An Interface is a named collection of typed functions.
//!
//! The Linker resolves imports against exports from other components,
//! producing a unified import table for the VM.

use std::collections::HashMap;

/// A typed function signature in an interface.
#[derive(Debug, Clone)]
pub struct FuncSig {
    pub name: String,
    pub params: Vec<ValType>,
    pub results: Vec<ValType>,
}

/// Value types for interface signatures.
#[derive(Debug, Clone, PartialEq)]
pub enum ValType {
    I32,
    I64,
    F64,
    String,
    Bool,
    List(Box<ValType>),
    Record(Vec<(String, ValType)>),
    /// Any — for dynamic languages that don't specify types
    Any,
}

/// An interface — a named collection of function signatures.
#[derive(Debug, Clone)]
pub struct Interface {
    pub name: String,
    pub functions: Vec<FuncSig>,
}

/// An export entry — maps interface function to implementation.
#[derive(Debug, Clone)]
pub enum ExportImpl {
    /// Host function index
    HostFn(usize),
    /// Bytecode chunk index
    ChunkFn(usize),
}

/// A component — a compiled module with import/export declarations.
#[derive(Debug, Clone)]
pub struct Component {
    /// Component name (from source file or vybe.toml)
    pub name: String,
    /// Source language
    pub language: Language,
    /// Compiled bytecode chunks
    pub chunks: Vec<crate::Chunk>,
    /// Import declarations: (interface_name, func_name)
    pub imports: Vec<(String, String)>,
    /// Export declarations: (interface_name, func_name) → implementation
    pub exports: HashMap<(String, String), ExportImpl>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Language {
    VB,
    JS,
}

/// The Component Model Linker.
/// Takes multiple components and resolves imports against exports.
pub struct Linker {
    /// All registered components
    components: Vec<Component>,
    /// Host-provided interfaces (wasi:*, vybe:*)
    host_exports: HashMap<(String, String), ExportImpl>,
}

impl Linker {
    pub fn new() -> Self {
        Linker {
            components: Vec::new(),
            host_exports: HashMap::new(),
        }
    }

    /// Register a host interface export.
    pub fn register_host_export(&mut self, interface: &str, func: &str, host_fn_idx: usize) {
        self.host_exports.insert(
            (interface.to_string(), func.to_string()),
            ExportImpl::HostFn(host_fn_idx),
        );
    }

    /// Register all host functions from a VM as exports.
    pub fn register_host_from_vm(&mut self, vm: &crate::VM) {
        for ((module, name), &idx) in &vm.host_registry {
            self.host_exports.insert(
                (module.clone(), name.clone()),
                ExportImpl::HostFn(idx),
            );
        }
    }

    /// Add a component to the linker.
    pub fn add_component(&mut self, component: Component) {
        self.components.push(component);
    }

    /// Link all components together.
    /// Returns merged chunks and a resolved import table.
    pub fn link(&self) -> Result<LinkResult, String> {
        // 1. Merge all chunks, adjusting chunk indices for each component
        let mut all_chunks: Vec<crate::Chunk> = Vec::new();
        let mut component_offsets: Vec<usize> = Vec::new();

        for comp in &self.components {
            component_offsets.push(all_chunks.len());
            all_chunks.extend(comp.chunks.clone());
        }

        // 2. Build export table from all components
        let mut all_exports: HashMap<(String, String), ExportImpl> = self.host_exports.clone();

        for (i, comp) in self.components.iter().enumerate() {
            let offset = component_offsets[i];
            for ((iface, func), impl_) in &comp.exports {
                let adjusted = match impl_ {
                    ExportImpl::ChunkFn(idx) => ExportImpl::ChunkFn(idx + offset),
                    ExportImpl::HostFn(idx) => ExportImpl::HostFn(*idx),
                };
                all_exports.insert((iface.clone(), func.clone()), adjusted);
            }
        }

        // 3. Resolve imports for each component
        let mut unresolved: Vec<(String, String, String)> = Vec::new(); // (component, iface, func)

        for comp in &self.components {
            for (iface, func) in &comp.imports {
                if !all_exports.contains_key(&(iface.clone(), func.clone())) {
                    unresolved.push((comp.name.clone(), iface.clone(), func.clone()));
                }
            }
        }

        if !unresolved.is_empty() {
            let msgs: Vec<String> = unresolved.iter()
                .map(|(c, i, f)| format!("  {}: {}:{}", c, i, f))
                .collect();
            return Err(format!("Unresolved imports:\n{}", msgs.join("\n")));
        }

        // 4. Build resolved import table for the merged chunks
        // The import table maps (module, name) → host_fn_idx or chunk_idx
        let mut resolved_imports: HashMap<(String, String), ExportImpl> = all_exports;

        Ok(LinkResult {
            chunks: all_chunks,
            exports: resolved_imports,
            component_offsets,
        })
    }
}

/// Result of linking multiple components.
pub struct LinkResult {
    /// All bytecode chunks, merged and adjusted
    pub chunks: Vec<crate::Chunk>,
    /// Resolved export table: (interface, func) → implementation
    pub exports: HashMap<(String, String), ExportImpl>,
    /// Offset of each component's chunks in the merged array
    pub component_offsets: Vec<usize>,
}
