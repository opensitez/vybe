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
use std::sync::{Mutex, OnceLock};

/// A binary-format loader registered by a platform crate — decodes a
/// module binary (.wasm, .class, …) into chunks. The VM itself ships no
/// codecs; platforms/* register theirs at host startup (e.g.
/// `vybe_platform_wasm::register()`), keyed by file extension.
pub type BinaryLoader = fn(&[u8]) -> Result<Vec<crate::Chunk>, String>;

static BINARY_LOADERS: OnceLock<Mutex<HashMap<String, BinaryLoader>>> = OnceLock::new();

/// Register a chunk-producing loader for a file extension (no dot,
/// lowercase — "wasm", "class"). Later registrations replace earlier
/// ones, so hosts can override a platform's default.
pub fn register_binary_loader(ext: &str, loader: BinaryLoader) {
    BINARY_LOADERS
        .get_or_init(|| Mutex::new(HashMap::new()))
        .lock()
        .unwrap()
        .insert(ext.to_lowercase(), loader);
}

fn binary_loader(ext: &str) -> Option<BinaryLoader> {
    BINARY_LOADERS.get()?.lock().unwrap().get(ext).copied()
}

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
    /// CM3 / WASI 0.3 future<T>
    Future(Box<ValType>),
    /// CM3 / WASI 0.3 stream<T>
    Stream(Box<ValType>),
    /// option<T> — used in stream read results and nullable returns
    Option(Box<ValType>),
    /// `result<T, E>`, `result<_, E>`, `result<T>` or bare `result` — EACH
    /// case's payload is optional.
    ///
    /// It used to demand both, which meant `result<_, error-code>` — the most
    /// common return shape in all of WASI 0.3.1, and tuple element 1 of every
    /// stream-producing function — could not be spelled at all. A caller had
    /// to invent a stand-in type for the absent `ok`, and that stand-in then
    /// fed `elem_size` and the payload offset, so the layout was wrong in a
    /// way nothing reported. Same class as the missing `Variant`: the layout
    /// code was right, the TYPE could not name the shape.
    Result(Option<Box<ValType>>, Option<Box<ValType>>),
    /// `variant` — the general N-case tagged union, each case optionally
    /// carrying a payload.
    ///
    /// `option` and `result` above are the two SPECIALISATIONS the Component
    /// Model despecialises into this, so they could be modelled without it.
    /// Nothing else could: `wasi:filesystem`'s `descriptor-type` is eight
    /// cases whose last is `other(option<string>)`, and there was no way to
    /// name that here at all. `canon_layout` already carried
    /// `alignment_variant`/`elem_size_variant` for exactly this shape — the
    /// layout half was written before the type existed.
    Variant(Vec<(String, Option<ValType>)>),
    /// Any — for dynamic languages that don't specify types
    Any,
    /// `own<T>` — an owned resource handle. The receiver takes ownership and is
    /// responsible for dropping it.
    ///
    /// Named by resource type (`"node"`, `"file-handle"`), not by index: a
    /// declaration binds to whatever is registered under that name, the same
    /// rule the type table already uses for a supertype it does not define.
    Own(String),
    /// `borrow<T>` — a borrowed resource handle. Valid for the duration of the
    /// call and never dropped by the callee.
    ///
    /// This is the ordinary shape for DOM operations: `append-child(borrow<node>,
    /// borrow<node>)` neither consumes its parent nor its child. Without the
    /// distinction every handle looks owned, and a host that dropped one would
    /// be indistinguishable from one that did not.
    Borrow(String),
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
    /// Type exports: types this component makes available.
    /// (interface_name, type_name) → TypeDef snapshot.
    pub type_exports: HashMap<(String, String), crate::TypeDef>,
    /// Type imports: types this component needs from other components.
    /// (interface_name, type_name)
    pub type_imports: Vec<(String, String)>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Language {
    VB,
    JS,
    CSharp,
    Python,
    Dart,
    Php,
    Ruby,
    Cobol,
    Wasm,
}

/// The Component Model Linker.
/// Takes multiple components and resolves imports against exports.
pub struct Linker {
    /// All registered components
    components: Vec<Component>,
    /// Host-provided interfaces (wasi:*, vybe:*)
    host_exports: HashMap<(String, String), ExportImpl>,
    /// Host-provided type exports sourced from Module Records.
    host_type_exports: HashMap<(String, String), crate::TypeDef>,
}

impl Linker {
    pub fn new() -> Self {
        Linker {
            components: Vec::new(),
            host_exports: HashMap::new(),
            host_type_exports: HashMap::new(),
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
        for (module, name, idx) in vm.iter_host_function_exports() {
            self.host_exports
                .insert((module, name), ExportImpl::HostFn(idx));
        }
        for (module, name, typedef) in vm.iter_host_type_exports() {
            self.host_type_exports.insert((module, name), typedef);
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

        // 3. Validate imports — each component's imports should be resolvable
        // from either component exports or host functions.
        // "*" module imports are resolved against all component exports.
        // Host imports (non-"*") are resolved at runtime by the VM.
        let mut unresolved: Vec<(String, String, String)> = Vec::new();

        for comp in &self.components {
            for (iface, func) in &comp.imports {
                if all_exports.contains_key(&(iface.clone(), func.clone())) {
                    continue; // resolved from component exports
                }
                if self
                    .host_exports
                    .contains_key(&(iface.clone(), func.clone()))
                {
                    continue; // resolved from host
                }
                // Non-"*" imports (e.g. "wasi:cli", "vybe:math") are host imports
                // resolved at runtime — don't flag as unresolved
                if iface != "*" {
                    continue;
                }
                // "*" import not found in any component export
                unresolved.push((comp.name.clone(), iface.clone(), func.clone()));
            }
        }

        if !unresolved.is_empty() {
            let msgs: Vec<String> = unresolved
                .iter()
                .map(|(c, i, f)| format!("  {}: {}:{}", c, i, f))
                .collect();
            return Err(format!("Unresolved imports:\n{}", msgs.join("\n")));
        }

        // 4. Resolve type imports/exports across components
        let mut all_type_exports: HashMap<(String, String), crate::TypeDef> =
            self.host_type_exports.clone();
        for comp in &self.components {
            for ((iface, name), typedef) in &comp.type_exports {
                all_type_exports.insert((iface.clone(), name.clone()), typedef.clone());
            }
        }

        // Check type imports are satisfied
        let mut unresolved_types: Vec<(String, String, String)> = Vec::new();
        for comp in &self.components {
            for (iface, name) in &comp.type_imports {
                if !all_type_exports.contains_key(&(iface.clone(), name.clone())) {
                    unresolved_types.push((comp.name.clone(), iface.clone(), name.clone()));
                }
            }
        }
        // Type imports that aren't resolved are warnings, not errors
        // (the type may exist in the host or be registered at runtime)

        // 5. CLS case resolution — rewrite lowercased property names in
        // case-insensitive components (VB) to match the canonical casing
        // from case-sensitive components (C#, JS, Python, Dart).
        //
        // Build canonical case map from all type fields and method names
        // across case-sensitive components. Then rewrite constant pool
        // entries in VB chunks where a lowercased string matches a
        // canonical-cased identifier.
        {
            let mut case_map: HashMap<String, String> = HashMap::new(); // lowercase → original

            // Collect canonical casing from case-sensitive components
            for comp in &self.components {
                // Skip case-insensitive languages — they're the ones we're fixing
                if comp.language == Language::VB || comp.language == Language::Cobol {
                    continue;
                }
                for chunk in &comp.chunks {
                    for entry in &chunk.types {
                        // Type name
                        case_map
                            .entry(entry.name.to_lowercase())
                            .or_insert_with(|| entry.name.clone());
                        // Fields
                        for field in &entry.fields {
                            case_map
                                .entry(field.to_lowercase())
                                .or_insert_with(|| field.clone());
                        }
                        // Methods
                        for (method, _) in &entry.methods {
                            case_map
                                .entry(method.to_lowercase())
                                .or_insert_with(|| method.clone());
                        }
                    }
                    // Also scan constant pools for string constants (property names used in code)
                    for val in &chunk.constants {
                        if let crate::Value::String(s) = val {
                            let lower = s.to_lowercase();
                            if lower != s.as_ref() {
                                // only if it HAS uppercase
                                case_map.entry(lower).or_insert_with(|| s.to_string());
                            }
                        }
                    }
                }
            }

            // Rewrite constant pools in case-insensitive component chunks (VB, COBOL)
            for (i, comp) in self.components.iter().enumerate() {
                if comp.language != Language::VB && comp.language != Language::Cobol {
                    continue;
                }
                let offset = component_offsets[i];
                for (ci, _chunk) in comp.chunks.iter().enumerate() {
                    let merged_ci = offset + ci;
                    let constants = &mut all_chunks[merged_ci].constants;
                    for val in constants.iter_mut() {
                        if let crate::Value::String(s) = val {
                            let lower = s.to_lowercase();
                            if lower == s.as_ref() {
                                // already lowercase (VB convention)
                                if let Some(canonical) = case_map.get(&lower) {
                                    if canonical.as_str() != s.as_ref() {
                                        *val = crate::Value::String(std::sync::Arc::from(
                                            canonical.as_str(),
                                        ));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // 6. Import unification — build unified import table and remap call_import indices.
        // Each component's script chunk (chunk 0) has its own imports. When merged,
        // call_import indices must refer to the unified table.
        {
            let mut unified_imports: Vec<crate::chunk::Import> = Vec::new();

            // For each component, build a mapping: old_import_idx → new_import_idx
            let mut remap_tables: Vec<Vec<u16>> = Vec::new();

            for comp in &self.components {
                let mut remap: Vec<u16> = Vec::new();
                // Component imports are on the script chunk (chunk 0)
                let comp_imports = if comp.chunks.is_empty() {
                    &[][..]
                } else {
                    &comp.chunks[0].imports[..]
                };
                for imp in comp_imports {
                    // Find or insert in unified table
                    let existing = unified_imports
                        .iter()
                        .position(|u| u.module == imp.module && u.name == imp.name);
                    let new_idx = if let Some(idx) = existing {
                        idx as u16
                    } else {
                        let idx = unified_imports.len() as u16;
                        unified_imports.push(imp.clone());
                        idx
                    };
                    remap.push(new_idx);
                }
                remap_tables.push(remap);
            }

            // Remap call_import operands in all merged chunks
            for (comp_idx, comp) in self.components.iter().enumerate() {
                let offset = component_offsets[comp_idx];
                let remap = &remap_tables[comp_idx];
                for ci in 0..comp.chunks.len() {
                    let merged_ci = offset + ci;
                    let code = &mut all_chunks[merged_ci].code;
                    let mut ip = 0;
                    while ip < code.len() {
                        if ip + 3 >= code.len() {
                            break;
                        }
                        let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                        if let Some(op) = crate::opcode::Op::decode(group, sub) {
                            if op == crate::opcode::Op::CALL {
                                // Remap import index: operand u16 is at ip+4..ip+5
                                if ip + 5 < code.len() {
                                    let old_idx =
                                        ((code[ip + 4] as u16) << 8) | (code[ip + 5] as u16);
                                    if (old_idx as usize) < remap.len() {
                                        let new_idx = remap[old_idx as usize];
                                        code[ip + 4] = (new_idx >> 8) as u8;
                                        code[ip + 5] = (new_idx & 0xff) as u8;
                                    }
                                }
                            }
                            let fmt = op.operand_format();
                            ip += 4;
                            ip += fmt.size_in(code, ip);
                        } else {
                            ip += 4;
                        }
                    }
                }
            }

            // Store unified imports on the first merged chunk (script chunk 0)
            if !all_chunks.is_empty() {
                all_chunks[0].imports = unified_imports;
            }
        }

        let resolved_exports: HashMap<(String, String), ExportImpl> = all_exports;

        // 7. Pre-resolve unified import table → ImportTarget for the VM.
        // This lets the VM skip runtime resolution for linked components.
        let mut resolved_imports = Vec::new();
        if !all_chunks.is_empty() {
            for imp in &all_chunks[0].imports {
                let key = (imp.module.clone(), imp.name.clone());
                if let Some(export) = resolved_exports.get(&key) {
                    match export {
                        ExportImpl::ChunkFn(ci) => {
                            let arity = if *ci < all_chunks.len() {
                                all_chunks[*ci].arity
                            } else {
                                0
                            };
                            resolved_imports.push(crate::vm::ImportTarget::ChunkFn {
                                chunk_index: *ci,
                                arity,
                            });
                        }
                        ExportImpl::HostFn(idx) => {
                            resolved_imports.push(crate::vm::ImportTarget::Host(*idx));
                        }
                    }
                } else if let Some(host_export) = self.host_exports.get(&key) {
                    match host_export {
                        ExportImpl::HostFn(idx) => {
                            resolved_imports.push(crate::vm::ImportTarget::Host(*idx));
                        }
                        _ => {
                            // Unresolved — will be resolved at runtime by the VM
                            resolved_imports
                                .push(crate::vm::ImportTarget::StdlibRedirect(imp.name.clone()));
                        }
                    }
                } else {
                    // Not resolved at link time — VM will resolve at runtime
                    resolved_imports
                        .push(crate::vm::ImportTarget::StdlibRedirect(imp.name.clone()));
                }
            }
        }

        Ok(LinkResult {
            chunks: all_chunks,
            exports: resolved_exports,
            component_offsets,
            type_exports: all_type_exports,
            resolved_imports,
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
    /// Resolved type exports: (interface, type_name) → TypeDef
    pub type_exports: HashMap<(String, String), crate::TypeDef>,
    /// Pre-resolved import table for the unified import list on chunks[0].
    /// Maps import index → ImportTarget. The VM can load this directly instead
    /// of resolving at runtime.
    pub resolved_imports: Vec<crate::vm::ImportTarget>,
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
            return Err(format!(
                "Import policy: absolute paths not allowed: {}",
                source
            ));
        }

        // Check parent traversal
        if !self.allow_parent_traversal && source.contains("..") {
            return Err(format!(
                "Import policy: parent directory traversal not allowed: {}",
                source
            ));
        }

        // Check extension
        let ext = std::path::Path::new(resolved_path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        if !self.allowed_extensions.iter().any(|e| *e == ext) {
            return Err(format!(
                "Import policy: extension '.{}' not allowed (allowed: {:?})",
                ext, self.allowed_extensions
            ));
        }

        // Check allowed directories
        if !self.allowed_dirs.is_empty() {
            let in_allowed = self
                .allowed_dirs
                .iter()
                .any(|d| resolved_path.starts_with(d.as_str()));
            if !in_allowed {
                return Err(format!(
                    "Import policy: path outside allowed directories: {}",
                    resolved_path
                ));
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
            return Err(format!(
                "Import policy: maximum module count ({}) reached",
                self.policy.max_modules
            ));
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
            _ if binary_loader(&ext).is_some() => {
                self.resolve_binary(&abs_path, binary_loader(&ext).unwrap())?
            }
            "wasm" => {
                return Err(format!(
                    ".wasm module resolution requires a registered loader (call vybe_platform_wasm::register() at startup): {}",
                    source
                ));
            }
            "js" => {
                return Err(format!(
                    "JS module resolution requires vybe_compiler_js (use .vybe project for cross-language): {}",
                    source
                ));
            }
            "vb" => {
                return Err(format!(
                    "VB module resolution requires vybe_compiler_vb (use .vybe project for cross-language): {}",
                    source
                ));
            }
            _ => return Err(format!("Unknown module type: {}", source)),
        };

        self.cache.insert(abs_path.clone(), module);
        Ok(&self.cache[&abs_path])
    }

    /// Resolve a binary module through a registered platform loader —
    /// read the file, decode to chunks, extract exports.
    fn resolve_binary(&self, path: &str, loader: BinaryLoader) -> Result<ResolvedModule, String> {
        let data = std::fs::read(path).map_err(|e| format!("Failed to read {}: {}", path, e))?;
        let chunks = loader(&data)?;

        let mut exports = HashMap::new();
        for (i, chunk) in chunks.iter().enumerate() {
            if chunk.name != "<script>" && !chunk.name.starts_with("func_") {
                exports.insert(
                    chunk.name.clone(),
                    ModuleExport::Function {
                        chunk_index: i,
                        arity: chunk.arity,
                    },
                );
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
