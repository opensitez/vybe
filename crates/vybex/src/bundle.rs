//! The universal compile unit. A `Bundle` represents everything needed to
//! compile and run — whether loaded from a single source file or a multi-file
//! project. The caller never cares which.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use crate::ast::*;
use crate::languages::Language;
use vybe_bytecode::{ExportEntry, ModuleRecord};

/// How the program starts.
#[derive(Debug, Clone)]
pub enum EntryPoint {
    /// Infer from code (Sub Main, main(), etc.)
    Auto,
    /// Launch a named form as the startup window.
    Form(String),
}

/// A source file within a bundle.
#[derive(Debug, Clone)]
pub struct SourceFile {
    pub path: PathBuf,
    pub code: String,
}

/// A pre-compiled WASM binary to link alongside source files.
#[derive(Debug, Clone)]
pub struct WasmFile {
    pub path: PathBuf,
    pub data: Vec<u8>,
}

/// Everything needed to compile and run.
pub struct Bundle {
    pub name: String,
    pub language: Language,
    pub sources: Vec<SourceFile>,
    pub wasm_files: Vec<WasmFile>,
    pub entry_point: EntryPoint,
}

/// What `Bundle::compile_full` returns — chunks + ESM import metadata
/// so the VM setup can install globals for read-as-value imports and
/// synthesize namespace objects for wildcard imports.
pub struct CompiledBundle {
    pub chunks: Vec<vybe_bytecode::Chunk>,
    pub host_imports: crate::compiler::HostImportMetadata,
}

impl Bundle {
    /// Parse all sources and compile to bytecode chunks.
    ///
    /// Legacy API — retains the `Vec<Chunk>` return shape for callers
    /// (tests, older code) that don't need the import metadata. Newer
    /// callers that install ESM host bindings should use
    /// [`Self::compile_full`].
    pub fn compile(&self) -> Result<Vec<vybe_bytecode::Chunk>, String> {
        self.compile_full().map(|r| r.chunks)
    }

    /// Compile and return chunks + ESM import metadata. Uses an empty
    /// module registry — no Adapter resolution. Call sites that need
    /// Adapter modules (`node:*` etc.) should use
    /// [`Self::compile_full_with_modules`].
    pub fn compile_full(&self) -> Result<CompiledBundle, String> {
        self.compile_full_with_modules(&std::collections::HashMap::new())
    }

    /// Compile with a read-only snapshot of `vm.modules` so the Linker
    /// can resolve Adapter-module imports (`import { X } from "node:http"`)
    /// by walking the re-export chain to the ultimate Synthetic target.
    ///
    /// The snapshot is flattened in this function — each adapter's
    /// `Indirect` exports are resolved transitively so the Compiler's
    /// Linker sees pre-resolved `(final_module, final_name)` pairs.
    pub fn compile_full_with_modules(
        &self,
        modules: &std::collections::HashMap<String, vybe_bytecode::ModuleRecord>,
    ) -> Result<CompiledBundle, String> {
        // Concatenate all sources
        let combined: String = self.sources.iter()
            .map(|s| s.code.as_str())
            .collect::<Vec<_>>()
            .join("\n");

        // Parse → common AST
        let mut module = (self.language.parse)(&combined)?;

        // Resolve imports relative to first source's directory
        if let Some(first) = self.sources.first() {
            let base_dir = first.path.parent().unwrap_or(Path::new("."));
            resolve_imports(&mut module, &self.language, base_dir);
        }

        // If the project starts with a form, inject startup AST:
        //   Dim __f = New FormName()
        //   Application.Run(__f)
        if let EntryPoint::Form(ref name) = self.entry_point {
            let new_expr = Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(name)),
                args: vec![],
            });
            module.body.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("__f".to_string()),
                    type_hint: None,
                    init: Some(new_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Dim,
            }));
            module.body.push(Statement::new(StmtKind::Expr(
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Application")),
                        field: "Run".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument {
                        value: Expression::ident("__f"),
                        name: None,
                        spread: false,
                        by_ref: false,
                    }],
                    optional: false,
                })
            )));
        }

        // Load profile + compile source code.
        //
        // Flatten the VM's module registry into a per-module map of
        // pre-resolved `(final_module, final_name)` pairs so the
        // compiler Linker can bind `import { X } from "node:http"` in
        // one lookup. Walks the `Indirect` chain from Adapter modules
        // through to a Synthetic `Function` export.
        let profile = crate::profile::parse_profile((self.language.profile_source)())?;
        let module_exports = flatten_module_exports(modules);
        let compile_result = crate::compiler::Compiler::with_profile(profile)
            .with_module_exports(module_exports)
            .compile_with_imports(&module)?;
        let mut chunks = compile_result.chunks;
        let host_imports = compile_result.host_imports;

        // Load and append WASM binary chunks
        for wf in &self.wasm_files {
            let wasm_chunks = vybe_bytecode::wasm::read_wasm(&wf.data)
                .map_err(|e| format!("WASM error in {}: {}", wf.path.display(), e))?;
            eprintln!("[vybex] Loaded {} chunks from {}", wasm_chunks.len(), wf.path.display());
            // Register WASM functions as globals so source code can call them
            for wc in &wasm_chunks {
                if !wc.name.is_empty() && wc.name != "<script>" {
                    // The WASM chunk index will be: current chunks.len() + position
                    eprintln!("  → fn {} (arity={})", wc.name, wc.arity);
                }
            }
            chunks.extend(wasm_chunks);
        }

        Ok(CompiledBundle { chunks, host_imports })
    }
}

/// Resolve `import { x } from "./file.js"` by parsing the imported file
/// and prepending its body to the main module.
fn resolve_imports(module: &mut Module, lang: &Language, base_dir: &Path) {
    let mut prepend: Vec<Statement> = Vec::new();
    for imp in &module.imports {
        let path_str = match &imp.kind {
            ImportKind::Named { path, .. } => path.clone(),
            ImportKind::Default { path, .. } => path.clone(),
            ImportKind::Simple { path, .. } => path.clone(),
            ImportKind::Wildcard { path, .. } => path.clone(),
        };
        // Host Component Model namespaces (`wasi:*`, `wasm:*`, `vybe:*`)
        // are not source files — the compiler binds them at compile time
        // from `module.imports` directly via `host_import_bindings`.
        // Skip filesystem resolution so we don't print spurious "no such
        // file" warnings.
        if path_str.starts_with("wasi:")
            || path_str.starts_with("wasm:")
            || path_str.starts_with("vybe:")
        {
            continue;
        }

        let resolved = base_dir.join(&path_str);
        if !should_resolve_source_import(&path_str, &resolved) {
            continue;
        }
        let source = match std::fs::read_to_string(&resolved) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("Warning: cannot resolve import '{}': {}", path_str, e);
                continue;
            }
        };
        let mut imported = match (lang.parse)(&source) {
            Ok(m) => m,
            Err(e) => {
                eprintln!("Warning: parse error in '{}': {}", path_str, e);
                continue;
            }
        };
        let import_dir = resolved.parent().unwrap_or(base_dir);
        resolve_imports(&mut imported, lang, import_dir);
        prepend.extend(imported.body);
    }
    prepend.append(&mut module.body);
    module.body = prepend;
}

fn should_resolve_source_import(path_str: &str, resolved: &Path) -> bool {
    if resolved.exists() {
        return true;
    }

    if path_str.starts_with('.') || path_str.starts_with('/') || path_str.starts_with('~') {
        return true;
    }

    if path_str.contains('/') || path_str.contains('\\') {
        return true;
    }

    matches!(
        resolved.extension().and_then(|ext| ext.to_str()).map(|ext| ext.to_ascii_lowercase()),
        Some(ext)
            if matches!(
                ext.as_str(),
                "vb" | "cs" | "js" | "ts" | "py" | "php" | "rb" | "dart" | "pas" | "cob" | "for" | "f90" | "wasm"
            )
    )
}

/// Flatten the VM's module registry so each exported name resolves to
/// a concrete `(module, func)` pair — walking the `Indirect` re-export
/// chain through Adapter modules to the ultimate Synthetic export.
///
/// ECMA-262 §16.2.1.6.2 `ResolveExport` done eagerly, once. The
/// compiler Linker then sees a flat `HashMap<specifier,
/// HashMap<name, (module, func)>>` and resolves user imports in one
/// lookup. Cycles (`export { A } from "m1"; export { A } from "m2"`
/// chained circular) are broken by a visit set; unresolved names drop
/// out of the map.
///
/// Public so tests and programmatic callers that bypass `Bundle` can
/// produce the same snapshot to feed into `Compiler::with_module_exports`.
pub fn flatten_module_exports(
    modules: &HashMap<String, ModuleRecord>,
) -> HashMap<String, HashMap<String, (String, String)>> {
    let mut out: HashMap<String, HashMap<String, (String, String)>> = HashMap::new();
    for (specifier, record) in modules {
        let mut resolved: HashMap<String, (String, String)> = HashMap::new();
        for (name, _) in &record.exports {
            let mut visited: Vec<(String, String)> = Vec::new();
            if let Some(target) = resolve_export(modules, specifier, name, &mut visited) {
                resolved.insert(name.clone(), target);
            }
        }
        if !resolved.is_empty() {
            out.insert(specifier.clone(), resolved);
        }
    }
    out
}

/// Validate that every import in the compiled chunks resolves to a
/// known target — ECMA-262 §16.2.1.6.2 `ResolveExport` applied
/// statically. Phase 8 of the ESM host-access migration: catch
/// unresolved imports at compile time so the runtime `setup_execution`
/// path only ever sees resolvable names.
///
/// Returns a list of diagnostic strings — each is a "module::name"
/// pair that couldn't be resolved. Empty list = fully resolved.
///
/// Known-exempt specifiers:
///   * `"*"`      — runtime wildcard dispatched via globals
///   * `"env"`    — WASM default env module (sin/cos polyfills, etc.)
/// Everything else must have a `ModuleRecord` entry with the given
/// name, following the same Indirect chain walk as
/// `flatten_module_exports`.
pub fn validate_imports_against_modules(
    chunks: &[vybe_bytecode::Chunk],
    modules: &HashMap<String, ModuleRecord>,
) -> Vec<String> {
    let mut unresolved = Vec::new();
    // Imports live on chunk[0] by convention.
    let imports_chunk = match chunks.first() {
        Some(c) => c,
        None => return unresolved,
    };
    for imp in &imports_chunk.imports {
        if imp.module == "*" || imp.module == "env" {
            continue;
        }
        // Name prefixed with `__vybe_` resolves via stdlib globals at
        // runtime — skip (stdlib emission guarantees the chunk is
        // registered).
        if imp.name.starts_with("__vybe_") {
            continue;
        }
        let Some(record) = modules.get(&imp.module) else {
            unresolved.push(format!("{}::{}", imp.module, imp.name));
            continue;
        };
        // Walk the Indirect chain — same resolver as the Phase 6
        // adapter flattener.
        let mut visited: Vec<(String, String)> = Vec::new();
        if resolve_export(modules, &imp.module, &imp.name, &mut visited).is_none() {
            let _ = record;
            unresolved.push(format!("{}::{}", imp.module, imp.name));
        }
    }
    unresolved
}

/// Recursive resolver — the `ResolveExport(exportName, resolveSet)`
/// abstract op from §16.2.1.6.2. Walks `Indirect` entries until it
/// hits a `Function` (the canonical terminal) or exhausts the chain.
/// `Value` / `ResourceType` exports aren't representable in the
/// `(module, func)` output shape yet; those names just drop out.
fn resolve_export(
    modules: &HashMap<String, ModuleRecord>,
    specifier: &str,
    name: &str,
    visited: &mut Vec<(String, String)>,
) -> Option<(String, String)> {
    let key = (specifier.to_string(), name.to_string());
    if visited.contains(&key) {
        // Circular re-export — bail. Per spec this is a SyntaxError;
        // MVP drops the binding and lets the runtime surface it if
        // the user actually tries to call it.
        return None;
    }
    visited.push(key);

    let record = modules.get(specifier)?;
    match record.exports.get(name)? {
        ExportEntry::Function { .. } => {
            // Terminal — Synthetic export. Bind to the module the
            // function is registered under, which is the specifier
            // that owns this record.
            Some((specifier.to_string(), name.to_string()))
        }
        ExportEntry::Indirect { from, name: src_name } => {
            resolve_export(modules, from, src_name, visited)
        }
        ExportEntry::Value(_) | ExportEntry::ResourceType { .. } => None,
    }
}
