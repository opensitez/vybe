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
    Wasm,
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
        let resolved_imports: HashMap<(String, String), ExportImpl> = all_exports;

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

// ============================================================
// ESM Integration — Source Phase Imports
// ============================================================

/// A resolved module with its exports ready for binding.
#[derive(Debug, Clone)]
pub struct ResolvedModule {
    /// Module source path (e.g., "./math.wasm", "./utils.js")
    pub source: String,
    /// Language of the module
    pub language: Language,
    /// Compiled bytecode chunks
    pub chunks: Vec<crate::Chunk>,
    /// Named exports: name → chunk_index (for functions)
    pub exports: HashMap<String, ModuleExport>,
}

#[derive(Debug, Clone)]
pub enum ModuleExport {
    /// A function export: chunk index + arity
    Function { chunk_index: usize, arity: u8 },
    /// A constant value export
    Value(crate::Value),
}

/// Security policy for module resolution.
/// Controls what files and paths ESM imports can access.
#[derive(Debug, Clone)]
pub struct ImportPolicy {
    /// Allowed file extensions (e.g., ["wasm", "js", "vb"])
    pub allowed_extensions: Vec<String>,
    /// Allowed directory prefixes (resolved paths must start with one of these).
    /// Empty = allow all. Use this to sandbox imports to the project directory.
    pub allowed_dirs: Vec<String>,
    /// Deny list: specific paths or patterns that are never allowed.
    pub denied_paths: Vec<String>,
    /// Maximum number of modules that can be imported (0 = unlimited).
    pub max_modules: usize,
    /// Whether to allow absolute paths (e.g., /usr/lib/math.wasm).
    /// Default: false — only relative paths allowed.
    pub allow_absolute_paths: bool,
    /// Whether to allow .. (parent directory traversal).
    /// Default: false — prevents escaping the project directory.
    pub allow_parent_traversal: bool,
}

impl ImportPolicy {
    /// Restrictive default: only .wasm, relative paths, no parent traversal.
    pub fn default() -> Self {
        ImportPolicy {
            allowed_extensions: vec!["wasm".into(), "js".into(), "vb".into()],
            allowed_dirs: Vec::new(),
            denied_paths: Vec::new(),
            max_modules: 64,
            allow_absolute_paths: false,
            allow_parent_traversal: false,
        }
    }

    /// No restrictions — for trusted environments (CLI with --unrestricted).
    pub fn unrestricted() -> Self {
        ImportPolicy {
            allowed_extensions: vec!["wasm".into(), "js".into(), "vb".into()],
            allowed_dirs: Vec::new(),
            denied_paths: Vec::new(),
            max_modules: 0,
            allow_absolute_paths: true,
            allow_parent_traversal: true,
        }
    }

    /// Sandbox to a single directory (for web/untrusted contexts).
    pub fn sandboxed(dir: impl Into<String>) -> Self {
        ImportPolicy {
            allowed_extensions: vec!["wasm".into()],
            allowed_dirs: vec![dir.into()],
            denied_paths: Vec::new(),
            max_modules: 16,
            allow_absolute_paths: false,
            allow_parent_traversal: false,
        }
    }

    /// Check if a resolved path is allowed by this policy.
    fn check(&self, resolved_path: &str, source: &str) -> Result<(), String> {
        // Check absolute paths
        if !self.allow_absolute_paths && (source.starts_with('/') || source.contains(":\\")) {
            return Err(format!("Import policy: absolute paths not allowed: {}", source));
        }

        // Check parent traversal
        if !self.allow_parent_traversal && source.contains("..") {
            return Err(format!("Import policy: parent directory traversal not allowed: {}", source));
        }

        // Check extension
        let ext = std::path::Path::new(resolved_path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        if !self.allowed_extensions.iter().any(|e| *e == ext) {
            return Err(format!("Import policy: extension '.{}' not allowed (allowed: {:?})", ext, self.allowed_extensions));
        }

        // Check allowed directories
        if !self.allowed_dirs.is_empty() {
            let in_allowed = self.allowed_dirs.iter().any(|d| resolved_path.starts_with(d.as_str()));
            if !in_allowed {
                return Err(format!("Import policy: path outside allowed directories: {}", resolved_path));
            }
        }

        // Check denied paths
        for denied in &self.denied_paths {
            if resolved_path.contains(denied.as_str()) {
                return Err(format!("Import policy: path denied: {}", resolved_path));
            }
        }

        Ok(())
    }
}

/// Resolves ESM-style imports at compile time.
/// Handles: .wasm, .js, .vb files.
/// Enforces ImportPolicy for security.
///
/// Usage:
///   import { add, multiply } from "./math.wasm"
///   import { format } from "./utils.js"
///   Imports "./helpers.vb"  (VB syntax)
pub struct ModuleResolver {
    /// Cache of already-resolved modules (by absolute path)
    pub cache: HashMap<String, ResolvedModule>,
    /// Base directory for relative path resolution
    pub base_dir: String,
    /// Security policy controlling what can be imported
    pub policy: ImportPolicy,
}

impl ModuleResolver {
    pub fn new(base_dir: impl Into<String>) -> Self {
        ModuleResolver {
            cache: HashMap::new(),
            base_dir: base_dir.into(),
            policy: ImportPolicy::default(),
        }
    }

    pub fn with_policy(base_dir: impl Into<String>, policy: ImportPolicy) -> Self {
        ModuleResolver {
            cache: HashMap::new(),
            base_dir: base_dir.into(),
            policy,
        }
    }

    /// Resolve a module source path, returning its exports.
    /// Enforces the import policy before loading.
    pub fn resolve(&mut self, source: &str) -> Result<&ResolvedModule, String> {
        // Check module count limit
        if self.policy.max_modules > 0 && self.cache.len() >= self.policy.max_modules {
            return Err(format!("Import policy: maximum module count ({}) reached", self.policy.max_modules));
        }

        let abs_path = self.resolve_path(source);

        // Enforce security policy
        self.policy.check(&abs_path, source)?;

        if self.cache.contains_key(&abs_path) {
            return Ok(&self.cache[&abs_path]);
        }

        let ext = std::path::Path::new(&abs_path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

        let module = match ext.as_str() {
            "wasm" => self.resolve_wasm(&abs_path)?,
            "js" => return Err(format!("JS module resolution requires vybe_compiler_js (use .vybe project for cross-language): {}", source)),
            "vb" => return Err(format!("VB module resolution requires vybe_compiler_vb (use .vybe project for cross-language): {}", source)),
            _ => return Err(format!("Unknown module type: {}", source)),
        };

        self.cache.insert(abs_path.clone(), module);
        Ok(&self.cache[&abs_path])
    }

    /// Resolve a .wasm module — read binary, extract exports.
    fn resolve_wasm(&self, path: &str) -> Result<ResolvedModule, String> {
        let data = std::fs::read(path)
            .map_err(|e| format!("Failed to read {}: {}", path, e))?;
        let chunks = crate::wasm::read_wasm(&data)?;

        let mut exports = HashMap::new();
        for (i, chunk) in chunks.iter().enumerate() {
            if chunk.name != "<script>" && !chunk.name.starts_with("func_") {
                exports.insert(chunk.name.clone(), ModuleExport::Function {
                    chunk_index: i,
                    arity: chunk.arity,
                });
            }
        }

        Ok(ResolvedModule {
            source: path.to_string(),
            language: Language::Wasm,
            chunks,
            exports,
        })
    }

    fn resolve_path(&self, source: &str) -> String {
        if source.starts_with('/') || source.contains(":\\") {
            source.to_string()
        } else {
            let base = std::path::Path::new(&self.base_dir);
            let resolved = base.join(source);
            resolved.to_string_lossy().to_string()
        }
    }
}
