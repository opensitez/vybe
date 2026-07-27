//! The universal compile unit. A `Bundle` represents everything needed to
//! compile and run — whether loaded from a single source file or a multi-file
//! project. The caller never cares which.

use crate::ast::*;
use crate::languages::Language;
use std::collections::{HashMap, HashSet, VecDeque};
use std::path::{Path, PathBuf};
use vybe_bytecode::{ExportEntry, ModuleRecord};

/// How the program starts.
#[derive(Debug, Clone)]
pub enum EntryPoint {
    /// Infer from code (Sub Main, main(), etc.)
    Auto,
    /// Launch a named form as the startup window.
    Form(String),
    /// Call a static method on a class as the program entry point, e.g. Program.Main().
    Method(String, String),
    /// Dynamic-eval in "completion" mode: the module's value is its last
    /// top-level expression. A language-agnostic feature (used by Python
    /// `eval`, JS `eval`, a REPL) — applied on the common AST, so it works
    /// for any front-end. See the `completion_value` eval attribute.
    EvalCompletion,
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
    /// Parse all sources into the common AST and apply the same
    /// bundle-level preparation that compilation uses: source-import
    /// resolution plus entry-point injection.
    ///
    /// Static-first rule: if an include/import can be resolved here,
    /// it should stay on the normal compiler path rather than falling
    /// through to runtime compilation. That keeps the resulting module
    /// graph as portable as possible across runtimes beyond Vybe.
    pub fn prepared_module(&self) -> Result<Module, String> {
        self.prepared_module_with_php_entry_override(None)
    }

    pub fn prepared_module_with_php_entry_override(
        &self,
        php_entry_path_override: Option<&Path>,
    ) -> Result<Module, String> {
        let (combined, php_blocks) = if self.language.name == "php" {
            let expanded = expand_php_bundle_sources_with_map_and_entry_path(
                &self.sources,
                php_entry_path_override,
            )?;
            let blocks = expanded.blocks.clone();
            (
                normalize_php_source_for_parser(&expanded.into_code()),
                Some(blocks),
            )
        } else if self.language.name == "cobol" {
            (
                self.sources
                    .iter()
                    .map(|s| rewrite_cobol_assign_paths(&s.code, &s.path))
                    .collect::<Vec<_>>()
                    .join("\n"),
                None,
            )
        } else {
            (
                self.sources
                    .iter()
                    .map(|s| s.code.as_str())
                    .collect::<Vec<_>>()
                    .join("\n"),
                None,
            )
        };

        if self.language.name == "php"
            && std::env::var_os("VYBEX_DEBUG_WRITE_EXPANDED_PHP").is_some()
        {
            let _ = std::fs::write("/tmp/vybex_expanded.php", &combined);
        }

        // Parse → common AST
        let mut module = (self.language.parse)(&combined).map_err(|err| {
            php_blocks
                .as_ref()
                .map(|blocks| annotate_php_parse_error(&err, blocks))
                .unwrap_or(err)
        })?;

        // Resolve imports relative to first source's directory
        if let Some(first) = self.sources.first() {
            let base_dir = first.path.parent().unwrap_or(Path::new("."));
            resolve_imports(&mut module, &self.language, base_dir);
        }

        // If the project starts with a static method (e.g. C# Program.Main()),
        // inject a call:  ClassName.MethodName()
        if let EntryPoint::Method(ref class_name, ref method_name) = self.entry_point {
            module
                .body
                .push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(class_name)),
                            field: method_name.clone(),
                            null_safe: false,
                        })),
                        args: vec![],
                        optional: false,
                    },
                ))));
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
            module
                .body
                .push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
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
                    },
                ))));
        }

        // Completion-value mode (the `completion_value` eval attribute):
        // the run's value is its last top-level expression. Rewrite the final
        // top-level expression statement into a `return` so the module yields
        // it. Purely on the common AST — language-agnostic.
        if matches!(self.entry_point, EntryPoint::EvalCompletion) {
            if let Some(last) = module.body.last_mut() {
                if let StmtKind::Expr(expr) = &last.kind {
                    let value = expr.clone();
                    last.kind = StmtKind::Return(Some(value));
                }
            }
        }

        Ok(module)
    }

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
        self.compile_full_with_modules_and_php_entry_override(
            &std::collections::HashMap::new(),
            None,
        )
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
        self.compile_full_with_modules_and_php_entry_override(modules, None)
    }

    pub fn compile_full_with_modules_and_php_entry_override(
        &self,
        modules: &std::collections::HashMap<String, vybe_bytecode::ModuleRecord>,
        php_entry_path_override: Option<&Path>,
    ) -> Result<CompiledBundle, String> {
        let module = self.prepared_module_with_php_entry_override(php_entry_path_override)?;
        self.compile_prepared_module(&module, modules)
    }

    /// Compile an already-prepared (possibly transformed) common-AST module with
    /// this bundle's full setup — profile, GUI namespace, module-export linking,
    /// and WASM chunk appending. Split out of `compile_full_with_modules` so the
    /// debugger's expression evaluator can transform the module (e.g. wrap the
    /// trailing expression in a `return`) and still reuse the exact same compile
    /// pipeline, language-agnostically.
    pub fn compile_prepared_module(
        &self,
        module: &Module,
        modules: &std::collections::HashMap<String, vybe_bytecode::ModuleRecord>,
    ) -> Result<CompiledBundle, String> {
        // Load profile + compile source code.
        //
        // Flatten the VM's module registry into a per-module map of
        // pre-resolved `(final_module, final_name)` pairs so the
        // compiler Linker can bind `import { X } from "node:http"` in
        // one lookup. Walks the `Indirect` chain from Adapter modules
        // through to a Synthetic `Function` export.
        // Seed the platform-specific namespace constants into the language SDK
        // once, so the generic profile parser (in `vybe_plugin`) stays free of any
        // platform reference. When platforms become loadable modules they will
        // register these themselves at load time.
        {
            static SEED: std::sync::Once = std::sync::Once::new();
            SEED.call_once(|| {
                let constants = vybe_bytecode::registry::platform_namespace_constants();
                let m = constants.iter().map(|c| **c).collect::<Vec<_>>()
                    .iter()
                    .map(|(n, v)| (n.to_string(), *v))
                    .collect();
                crate::profile::register_dotnet_namespace_constants(m);
            });
        }
        let mut profile = crate::profile::parse_profile((self.language.profile_source)())?;

        // Add shared GUI namespace automatically to all languages
        // (replaces per-language profile duplication)
        add_shared_gui_namespace(&mut profile);

        let module_exports = flatten_module_exports(modules);
        let value_exports = flatten_module_value_exports(modules);
        let compile_result = crate::compiler::Compiler::with_profile(profile)
            .with_module_exports(module_exports)
            .with_module_value_exports(value_exports)
            .compile_with_imports(module)?;
        let mut chunks = compile_result.chunks;
        let host_imports = compile_result.host_imports;

        // Load and append WASM binary chunks
        for wf in &self.wasm_files {
            // Through the registry: the platform that can decode this format
            // registered a reader. The compiler does not name `vybe_platform_wasm`.
            let wasm_chunks = vybe_bytecode::registry::platform_read_binary_module(&wf.data)
                .ok_or_else(|| {
                    format!(
                        "no platform registered a reader for {}",
                        wf.path.display()
                    )
                })?
                .map_err(|e| format!("WASM error in {}: {}", wf.path.display(), e))?;
            if std::env::var_os("VYBE_TRACE").is_some() {
                eprintln!(
                    "[vybex] Loaded {} chunks from {}",
                    wasm_chunks.len(),
                    wf.path.display()
                );
                // The WASM chunk index will be: current chunks.len() + position
                for wc in &wasm_chunks {
                    if !wc.name.is_empty() && wc.name != "<script>" {
                        eprintln!("  → fn {} (arity={})", wc.name, wc.arity);
                    }
                }
            }
            chunks.extend(wasm_chunks);
        }

        Ok(CompiledBundle {
            chunks,
            host_imports,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PhpIncludeKind {
    Include,
    IncludeOnce,
    Require,
    RequireOnce,
}

#[derive(Debug, Clone)]
struct ExpandedPhpBlock {
    path: PathBuf,
    start_line: usize,
    end_line: usize,
}

#[derive(Debug, Clone, Default)]
struct ExpandedPhpSource {
    lines: Vec<String>,
    blocks: Vec<ExpandedPhpBlock>,
}

impl ExpandedPhpSource {
    fn into_code(self) -> String {
        self.lines.join("\n")
    }
}

fn record_expanded_line(expanded: &mut ExpandedPhpSource, path: &Path, line: String) {
    expanded.lines.push(line);
    let line_no = expanded.lines.len();
    if let Some(last) = expanded.blocks.last_mut()
        && last.path == path
        && last.end_line + 1 == line_no
    {
        last.end_line = line_no;
        return;
    }
    expanded.blocks.push(ExpandedPhpBlock {
        path: path.to_path_buf(),
        start_line: line_no,
        end_line: line_no,
    });
}

fn merge_expanded_block(blocks: &mut Vec<ExpandedPhpBlock>, block: ExpandedPhpBlock) {
    if let Some(last) = blocks.last_mut()
        && last.path == block.path
        && last.end_line + 1 == block.start_line
    {
        last.end_line = block.end_line;
        return;
    }
    blocks.push(block);
}

fn append_expanded_source(expanded: &mut ExpandedPhpSource, nested: ExpandedPhpSource) {
    let line_offset = expanded.lines.len();
    expanded.lines.extend(nested.lines);
    for mut block in nested.blocks {
        block.start_line += line_offset;
        block.end_line += line_offset;
        merge_expanded_block(&mut expanded.blocks, block);
    }
}

fn extract_php_parse_error_line(err: &str) -> Option<usize> {
    err.lines().find_map(|line| {
        let rest = line.split_once("-->")?.1.trim();
        rest.split(':').next()?.trim().parse::<usize>().ok()
    })
}

fn annotate_php_parse_error(err: &str, blocks: &[ExpandedPhpBlock]) -> String {
    let Some(expanded_line) = extract_php_parse_error_line(err) else {
        return err.to_string();
    };
    let Some(block) = blocks
        .iter()
        .rev()
        .find(|block| expanded_line >= block.start_line && expanded_line <= block.end_line)
    else {
        return err.to_string();
    };
    let source_line = expanded_line - block.start_line + 1;
    format!("{}\nsource: {}:{}", err, block.path.display(), source_line,)
}

#[allow(dead_code)]
fn expand_php_bundle_sources(sources: &[SourceFile]) -> Result<String, String> {
    Ok(expand_php_bundle_sources_with_map(sources)?.into_code())
}

#[allow(dead_code)]
fn expand_php_bundle_sources_with_map(sources: &[SourceFile]) -> Result<ExpandedPhpSource, String> {
    expand_php_bundle_sources_with_map_and_entry_path(sources, None)
}

fn expand_php_bundle_sources_with_map_and_entry_path(
    sources: &[SourceFile],
    entry_path_override: Option<&Path>,
) -> Result<ExpandedPhpSource, String> {
    let mut included_once = HashSet::new();
    let mut constants: HashMap<String, String> = HashMap::new();
    let mut expanded = ExpandedPhpSource::default();
    for source in sources {
        let entry_path = entry_path_override.unwrap_or(&source.path);
        append_expanded_source(
            &mut expanded,
            expand_php_source_file(
                entry_path,
                &source.path,
                &source.code,
                &mut included_once,
                &mut constants,
            )?,
        );
    }
    Ok(expanded)
}

fn expand_php_source_file(
    entry_path: &Path,
    source_path: &Path,
    code: &str,
    included_once: &mut HashSet<PathBuf>,
    constants: &mut HashMap<String, String>,
) -> Result<ExpandedPhpSource, String> {
    let rewritten_code = rewrite_php_magic_constants(code, source_path);
    let mut aliases: HashMap<String, String> = HashMap::new();
    let mut recent_optional_guards: VecDeque<PathBuf> = VecDeque::new();
    let mut recent_optional_plugin_dirs: VecDeque<PathBuf> = VecDeque::new();
    let mut brace_depth = 0usize;
    let mut previous_nonempty_line: Option<String> = None;
    let mut out = ExpandedPhpSource::default();

    let mut lines = rewritten_code.lines().peekable();
    while let Some(line) = lines.next() {
        let trimmed = line.trim();

        for guarded in extract_php_file_exists_guards(trimmed, source_path, &aliases, constants) {
            recent_optional_guards.push_back(guarded);
            if recent_optional_guards.len() > 8 {
                recent_optional_guards.pop_front();
            }
        }
        for guarded in extract_php_plugin_active_guards(trimmed, source_path, &aliases, constants) {
            recent_optional_plugin_dirs.push_back(guarded);
            if recent_optional_plugin_dirs.len() > 8 {
                recent_optional_plugin_dirs.pop_front();
            }
        }

        if brace_depth == 0
            && let Some((name, value)) =
                parse_php_alias_assignment(trimmed, source_path, &aliases, constants)
        {
            aliases.insert(name, value);
            record_expanded_line(&mut out, source_path, line.to_string());
            brace_depth = update_php_brace_depth(brace_depth, line);
            continue;
        }

        if let Some((name, value)) =
            parse_php_constant_assignment(trimmed, source_path, &aliases, constants)
        {
            constants.insert(name, value);
            record_expanded_line(&mut out, source_path, line.to_string());
            brace_depth = update_php_brace_depth(brace_depth, line);
            continue;
        }

        if let Some((kind, expr)) = parse_php_include_statement(trimmed) {
            if previous_nonempty_line
                .as_deref()
                .and_then(parse_positive_defined_guard)
                .is_some_and(|name| !constants.contains_key(&name))
            {
                brace_depth = update_php_brace_depth(brace_depth, line);
                previous_nonempty_line =
                    next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                continue;
            }
            if brace_depth > 0 {
                record_expanded_line(&mut out, source_path, line.to_string());
                brace_depth = update_php_brace_depth(brace_depth, line);
                previous_nonempty_line =
                    next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                continue;
            }
            let Some(resolved) = eval_php_path_expression(expr, source_path, &aliases, constants)
            else {
                if is_dynamic_php_include_expression(expr, constants) {
                    brace_depth = update_php_brace_depth(brace_depth, line);
                    previous_nonempty_line =
                        next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                    continue;
                }
                return Err(format!(
                    "Unsupported PHP include path expression in {}: {}",
                    source_path.display(),
                    trimmed,
                ));
            };
            let include_path =
                resolve_php_include_path_with_entry(entry_path, source_path, &resolved);
            let canonical = std::fs::canonicalize(&include_path).unwrap_or(include_path.clone());

            let should_expand = match kind {
                PhpIncludeKind::IncludeOnce | PhpIncludeKind::RequireOnce => {
                    included_once.insert(canonical.clone())
                }
                PhpIncludeKind::Include | PhpIncludeKind::Require => true,
            };

            if should_expand {
                let include_source = match std::fs::read_to_string(&canonical) {
                    Ok(source) => source,
                    Err(err) => {
                        if err.kind() == std::io::ErrorKind::NotFound
                            && recent_optional_guards
                                .iter()
                                .any(|guarded| guarded == &canonical)
                        {
                            continue;
                        }
                        if err.kind() == std::io::ErrorKind::NotFound
                            && recent_optional_plugin_dirs
                                .iter()
                                .any(|guarded| canonical.starts_with(guarded))
                        {
                            continue;
                        }
                        return Err(format!(
                            "PHP include error reading {}: {}",
                            canonical.display(),
                            err
                        ));
                    }
                };

                if brace_depth > 0 && php_include_contains_inline_html(&include_source) {
                    record_expanded_line(&mut out, source_path, line.to_string());
                    brace_depth = update_php_brace_depth(brace_depth, line);
                    previous_nonempty_line =
                        next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                    continue;
                }

                if php_include_contains_top_level_return(&include_source) {
                    record_expanded_line(&mut out, source_path, line.to_string());
                    brace_depth = update_php_brace_depth(brace_depth, line);
                    previous_nonempty_line =
                        next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                    continue;
                }

                let normalized_include = normalize_php_include_for_inlining(&include_source);
                append_expanded_source(
                    &mut out,
                    expand_php_source_file(
                        entry_path,
                        &canonical,
                        &normalized_include,
                        included_once,
                        constants,
                    )?,
                );
            }
            brace_depth = update_php_brace_depth(brace_depth, line);
            continue;
        }

        if starts_multiline_php_alternative_control_header(line) {
            let mut header_lines = vec![line.to_string()];
            let mut normalized_block = None;
            let mut scanned_lines = 0usize;

            while let Some(next_line) = lines.peek().copied() {
                if should_stop_multiline_php_header_scan(next_line) || scanned_lines >= 12 {
                    break;
                }
                header_lines.push(next_line.to_string());
                lines.next();
                scanned_lines += 1;
                let candidate = header_lines.join("\n");
                if let Some(normalized) =
                    normalize_multiline_php_alternative_control_syntax(&candidate)
                {
                    normalized_block = Some(normalized);
                    break;
                }
            }

            let expanded_lines = normalized_block.unwrap_or_else(|| header_lines.join("\n"));
            for expanded_line in expanded_lines.lines() {
                record_expanded_line(&mut out, source_path, expanded_line.to_string());
            }
            for original_line in &header_lines {
                brace_depth = update_php_brace_depth(brace_depth, original_line);
                previous_nonempty_line =
                    next_previous_nonempty_line(original_line.trim(), &previous_nonempty_line);
            }
            continue;
        }

        record_expanded_line(
            &mut out,
            source_path,
            normalize_php_alternative_control_syntax(line),
        );
        brace_depth = update_php_brace_depth(brace_depth, line);
        previous_nonempty_line = next_previous_nonempty_line(trimmed, &previous_nonempty_line);
    }

    Ok(out)
}

fn absolutize_path(path: &Path) -> PathBuf {
    if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(path)
    }
}

fn rewrite_cobol_assign_paths(source: &str, source_path: &Path) -> String {
    if !source.to_ascii_uppercase().contains("ASSIGN TO") {
        return source.to_string();
    }

    let base_dir = absolutize_path(source_path)
        .parent()
        .unwrap_or(Path::new("."))
        .to_path_buf();

    source
        .split_inclusive('\n')
        .map(|line| rewrite_cobol_assign_path_line(line, &base_dir))
        .collect()
}

fn rewrite_cobol_assign_path_line(line: &str, base_dir: &Path) -> String {
    let uppercase = line.to_ascii_uppercase();
    let Some(assign_idx) = uppercase.find("ASSIGN TO") else {
        return line.to_string();
    };

    let search_start = assign_idx + "ASSIGN TO".len();
    let Some((quote_offset, quote)) = line[search_start..]
        .char_indices()
        .find(|(_, ch)| *ch == '\'' || *ch == '"')
    else {
        return line.to_string();
    };

    let literal_start = search_start + quote_offset + quote.len_utf8();
    let Some(literal_end_rel) = line[literal_start..].find(quote) else {
        return line.to_string();
    };
    let literal_end = literal_start + literal_end_rel;
    let raw_path = &line[literal_start..literal_end];
    let candidate = Path::new(raw_path);
    if candidate.is_absolute() || raw_path.contains("://") {
        return line.to_string();
    }

    let resolved = base_dir.join(candidate);
    let mut rewritten = String::with_capacity(line.len() + resolved.as_os_str().len());
    rewritten.push_str(&line[..literal_start]);
    rewritten.push_str(&resolved.to_string_lossy());
    rewritten.push_str(&line[literal_end..]);
    rewritten
}

fn resolve_php_include_path(source_path: &Path, resolved: &str) -> PathBuf {
    resolve_php_include_path_with_entry(source_path, source_path, resolved)
}

fn resolve_php_include_path_with_entry(
    entry_path: &Path,
    source_path: &Path,
    resolved: &str,
) -> PathBuf {
    let candidate = PathBuf::from(resolved);
    if candidate.is_absolute() {
        candidate
    } else if is_explicit_relative_php_include(&candidate) {
        absolutize_path(source_path)
            .parent()
            .unwrap_or(Path::new("."))
            .join(candidate)
    } else {
        absolutize_path(entry_path)
            .parent()
            .unwrap_or(Path::new("."))
            .join(candidate)
    }
}

fn is_explicit_relative_php_include(path: &Path) -> bool {
    let raw = path.to_string_lossy();
    raw == "."
        || raw == ".."
        || raw.starts_with("./")
        || raw.starts_with("../")
        || raw.starts_with(".\\")
        || raw.starts_with("..\\")
}

fn parse_php_alias_assignment(
    line: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Option<(String, String)> {
    let statement = php_statement_prefix(line)?;
    if !statement.starts_with('$') {
        return None;
    }
    let (lhs, rhs) = statement.split_once('=')?;
    let name = lhs.trim().strip_prefix('$')?.trim();
    if name.is_empty()
        || !name
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        return None;
    }
    let value = eval_php_path_expression(rhs.trim(), source_path, aliases, constants)?;
    Some((name.to_string(), value))
}

fn parse_php_constant_assignment(
    line: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Option<(String, String)> {
    let statement = php_statement_prefix(line)?;

    if let Some(args) = statement.strip_prefix("define").map(str::trim) {
        let args = args.strip_prefix('(')?.strip_suffix(')')?.trim();
        let parts = split_php_function_args(args);
        if parts.len() < 2 {
            return None;
        }
        let name = parse_php_constant_name(parts[0].trim())?;
        let value = eval_php_path_expression(parts[1].trim(), source_path, aliases, constants)?;
        return Some((name, value));
    }

    let declaration = statement.strip_prefix("const ")?.trim();
    let (name, expr) = declaration.split_once('=')?;
    let name = name.trim();
    if name.is_empty()
        || !name
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        return None;
    }
    let value = eval_php_path_expression(expr.trim(), source_path, aliases, constants)?;
    Some((name.to_string(), value))
}

fn parse_php_constant_name(expr: &str) -> Option<String> {
    let expr = expr.trim();
    expr.strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .or_else(|| expr.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')))
        .map(|name| name.to_string())
}

fn parse_php_include_statement(line: &str) -> Option<(PhpIncludeKind, &str)> {
    let statement = php_statement_prefix(line)?;
    let kinds = [
        ("require_once", PhpIncludeKind::RequireOnce),
        ("include_once", PhpIncludeKind::IncludeOnce),
        ("require", PhpIncludeKind::Require),
        ("include", PhpIncludeKind::Include),
    ];

    for (kw, kind) in kinds {
        if let Some(rest) = statement.strip_prefix(kw) {
            if rest.starts_with(|ch: char| ch.is_ascii_alphanumeric() || ch == '_') {
                continue;
            }
            let mut expr = rest.trim();
            if expr.starts_with('(') && expr.ends_with(')') {
                expr = expr[1..expr.len() - 1].trim();
            }
            return Some((kind, expr));
        }
    }
    None
}

fn php_statement_prefix(line: &str) -> Option<&str> {
    let mut quote: Option<char> = None;
    let chars: Vec<(usize, char)> = line.char_indices().collect();
    let mut index = 0usize;

    while index < chars.len() {
        let (byte_idx, ch) = chars[index];
        if let Some(q) = quote {
            if ch == '\\' {
                index += 2;
                continue;
            }
            if ch == q {
                quote = None;
            }
            index += 1;
            continue;
        }

        match ch {
            '\'' | '"' => quote = Some(ch),
            ';' => return Some(line[..byte_idx].trim()),
            _ => {}
        }
        index += 1;
    }

    None
}

fn is_dynamic_php_include_expression(expr: &str, constants: &HashMap<String, String>) -> bool {
    let expr = expr.trim();
    let expr = expr
        .strip_prefix('(')
        .and_then(|inner| inner.strip_suffix(')'))
        .map(str::trim)
        .unwrap_or(expr);
    let parts = split_php_concat(expr);
    !parts.is_empty()
        && parts
            .into_iter()
            .all(|part| is_dynamic_php_path_atom(part.trim(), constants))
}

fn is_dynamic_php_path_atom(atom: &str, constants: &HashMap<String, String>) -> bool {
    let atom = atom.trim();
    if atom.is_empty() {
        return true;
    }
    if atom
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .is_some()
    {
        return true;
    }
    if atom
        .strip_prefix('\'')
        .and_then(|s| s.strip_suffix('\''))
        .is_some()
    {
        return true;
    }
    if atom.eq("__DIR__") || atom.eq("__FILE__") {
        return true;
    }
    if atom.starts_with('$') {
        return true;
    }
    if atom.starts_with("dirname(") && atom.ends_with(')') {
        let inner = &atom["dirname(".len()..atom.len() - 1];
        return is_dynamic_php_include_expression(inner, constants);
    }
    if atom.ends_with(')') && atom.contains('(') {
        return true;
    }
    if atom
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        return true;
    }
    false
}

fn extract_php_file_exists_guards(
    line: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Vec<PathBuf> {
    let mut guards = Vec::new();
    let mut search = line;

    while let Some(index) = search.find("file_exists(") {
        let after = &search[index + "file_exists(".len()..];
        let mut depth = 1usize;
        let mut quote: Option<char> = None;
        let mut end = None;
        let chars: Vec<(usize, char)> = after.char_indices().collect();
        let mut pos = 0usize;

        while pos < chars.len() {
            let (byte_idx, ch) = chars[pos];
            if let Some(q) = quote {
                if ch == '\\' {
                    pos += 2;
                    continue;
                }
                if ch == q {
                    quote = None;
                }
                pos += 1;
                continue;
            }

            match ch {
                '\'' | '"' => quote = Some(ch),
                '(' => depth += 1,
                ')' => {
                    depth = depth.saturating_sub(1);
                    if depth == 0 {
                        end = Some(byte_idx);
                        break;
                    }
                }
                _ => {}
            }
            pos += 1;
        }

        let Some(end_idx) = end else {
            break;
        };
        let expr = after[..end_idx].trim();
        if let Some(resolved) = eval_php_path_expression(expr, source_path, aliases, constants) {
            let include_path = resolve_php_include_path(source_path, &resolved);
            let canonical = std::fs::canonicalize(&include_path).unwrap_or(include_path);
            guards.push(canonical);
        }
        search = &after[end_idx + 1..];
    }

    guards
}

fn extract_php_plugin_active_guards(
    line: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Vec<PathBuf> {
    let line = line.trim();
    if (!line.starts_with("if") && !line.starts_with("elseif"))
        || line.contains("! is_plugin_active")
    {
        return Vec::new();
    }

    let Some(start) = line.find("is_plugin_active(") else {
        return Vec::new();
    };
    let after = &line[start + "is_plugin_active(".len()..];
    let Some(end) = after.find(')') else {
        return Vec::new();
    };

    let arg = after[..end].trim();
    let Some(plugin_rel) = eval_php_path_expression(arg, source_path, aliases, constants) else {
        return Vec::new();
    };
    let Some(plugin_root) = constants.get("WP_PLUGIN_DIR") else {
        return Vec::new();
    };

    let plugin_path =
        PathBuf::from(plugin_root.trim_end_matches('/')).join(plugin_rel.trim_start_matches('/'));
    let plugin_dir = plugin_path.parent().unwrap_or(&plugin_path).to_path_buf();
    vec![std::fs::canonicalize(&plugin_dir).unwrap_or(plugin_dir)]
}

fn update_php_brace_depth(mut depth: usize, line: &str) -> usize {
    let mut quote: Option<char> = None;
    let chars: Vec<char> = line.chars().collect();
    let mut index = 0usize;

    while index < chars.len() {
        let ch = chars[index];
        if let Some(q) = quote {
            if ch == '\\' {
                index += 2;
                continue;
            }
            if ch == q {
                quote = None;
            }
            index += 1;
            continue;
        }

        match ch {
            '\'' | '"' => quote = Some(ch),
            '{' => depth += 1,
            '}' => depth = depth.saturating_sub(1),
            _ => {}
        }
        index += 1;
    }

    depth
}

fn next_previous_nonempty_line<'a>(
    trimmed: &'a str,
    current: &'a Option<String>,
) -> Option<String> {
    if trimmed.is_empty() {
        current.clone()
    } else {
        Some(trimmed.to_string())
    }
}

fn parse_positive_defined_guard(line: &str) -> Option<String> {
    let line = line.trim();
    if !line.starts_with("if") && !line.starts_with("elseif") {
        return None;
    }
    if line.contains("! defined") {
        return None;
    }
    let start = line.find("defined(")? + "defined(".len();
    let after = &line[start..];
    let end = after.find(')')?;
    parse_php_constant_name(after[..end].trim())
}

fn eval_php_path_expression(
    expr: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Option<String> {
    let mut out = String::new();
    for part in split_php_concat(expr) {
        out.push_str(&eval_php_path_atom(
            part.trim(),
            source_path,
            aliases,
            constants,
        )?);
    }
    Some(out)
}

fn split_php_function_args(args: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut depth = 0usize;
    let mut quote: Option<char> = None;
    let chars: Vec<(usize, char)> = args.char_indices().collect();
    let mut index = 0usize;

    while index < chars.len() {
        let (byte_idx, ch) = chars[index];
        if let Some(q) = quote {
            if ch == '\\' {
                index += 2;
                continue;
            }
            if ch == q {
                quote = None;
            }
            index += 1;
            continue;
        }

        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => depth += 1,
            ')' => depth = depth.saturating_sub(1),
            ',' if depth == 0 => {
                parts.push(args[start..byte_idx].trim());
                start = byte_idx + ch.len_utf8();
            }
            _ => {}
        }
        index += 1;
    }

    let tail = args[start..].trim();
    if !tail.is_empty() {
        parts.push(tail);
    }
    parts
}

fn split_php_concat(expr: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut depth = 0usize;
    let mut quote: Option<char> = None;
    let chars: Vec<(usize, char)> = expr.char_indices().collect();
    let mut index = 0usize;

    while index < chars.len() {
        let (byte_idx, ch) = chars[index];
        if let Some(q) = quote {
            if ch == '\\' {
                index += 2;
                continue;
            }
            if ch == q {
                quote = None;
            }
            index += 1;
            continue;
        }

        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => depth += 1,
            ')' => depth = depth.saturating_sub(1),
            '.' if depth == 0 => {
                parts.push(expr[start..byte_idx].trim());
                start = byte_idx + ch.len_utf8();
            }
            _ => {}
        }
        index += 1;
    }

    let tail = expr[start..].trim();
    if !tail.is_empty() {
        parts.push(tail);
    }
    parts
}

fn eval_php_path_atom(
    atom: &str,
    source_path: &Path,
    aliases: &HashMap<String, String>,
    constants: &HashMap<String, String>,
) -> Option<String> {
    let atom = atom.trim();
    if atom.is_empty() {
        return Some(String::new());
    }

    if let Some(stripped) = atom.strip_prefix('"').and_then(|s| s.strip_suffix('"')) {
        if php_double_quoted_string_has_interpolation(stripped) {
            return None;
        }
        return Some(stripped.to_string());
    }
    if let Some(stripped) = atom.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')) {
        return Some(stripped.to_string());
    }
    if atom.eq("__DIR__") {
        return Some(
            absolutize_path(source_path)
                .parent()
                .unwrap_or(Path::new("."))
                .to_string_lossy()
                .into_owned(),
        );
    }
    if atom.eq("__FILE__") {
        return Some(absolutize_path(source_path).to_string_lossy().into_owned());
    }
    if let Some(name) = atom.strip_prefix('$') {
        return aliases.get(name).cloned();
    }
    if atom
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        if let Some(value) = constants.get(atom) {
            return Some(value.clone());
        }
    }
    if atom.starts_with("dirname(") && atom.ends_with(')') {
        let inner = &atom["dirname(".len()..atom.len() - 1];
        let resolved = eval_php_path_expression(inner, source_path, aliases, constants)?;
        let parent = Path::new(&resolved).parent().unwrap_or(Path::new("."));
        return Some(parent.to_string_lossy().into_owned());
    }
    None
}

fn php_double_quoted_string_has_interpolation(content: &str) -> bool {
    let bytes = content.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'\\' {
            index += 2;
            continue;
        }
        if bytes[index] == b'$' {
            return true;
        }
        index += 1;
    }
    false
}

fn normalize_php_include_for_inlining(source: &str) -> String {
    let mut out = String::new();
    let mut in_php = true;
    let segments = split_mixed_php_include_source(source);
    let last_index = segments.len().saturating_sub(1);

    for (index, segment) in segments.into_iter().enumerate() {
        match segment {
            MixedPhpIncludeSegment::Html(html) => {
                if html.is_empty()
                    || (html.trim().is_empty() && (index == 0 || index == last_index))
                {
                    continue;
                }
                if in_php {
                    out.push_str("?>");
                    in_php = false;
                }
                out.push_str(html);
            }
            MixedPhpIncludeSegment::Echo {
                expr,
                has_close_tag,
            } => {
                let expr = expr.trim();
                if !expr.is_empty() {
                    if !in_php {
                        out.push_str("<?php ");
                        in_php = true;
                    }
                    out.push_str("echo ");
                    out.push_str(expr);
                    out.push_str(";\n");
                }
                if has_close_tag {
                    if in_php {
                        out.push_str("?>");
                        in_php = false;
                    }
                }
            }
            MixedPhpIncludeSegment::Code {
                code,
                has_close_tag,
            } => {
                if !in_php {
                    out.push_str("<?php");
                    in_php = true;
                }
                out.push_str(code);
                if has_close_tag && php_code_block_needs_terminator(code) {
                    out.push(';');
                }
                if has_close_tag {
                    out.push_str("?>");
                    in_php = false;
                } else if !out.ends_with('\n') {
                    out.push('\n');
                }
            }
        }
    }

    if !in_php {
        out.push_str("<?php\n");
    }

    out
}

fn php_include_contains_inline_html(source: &str) -> bool {
    split_mixed_php_include_source(source)
        .into_iter()
        .any(|segment| match segment {
            MixedPhpIncludeSegment::Html(html) => !html.trim().is_empty(),
            _ => false,
        })
}

fn php_include_contains_top_level_return(source: &str) -> bool {
    split_mixed_php_include_source(source)
        .into_iter()
        .any(|segment| match segment {
            MixedPhpIncludeSegment::Code { code, .. }
            | MixedPhpIncludeSegment::Echo { expr: code, .. } => {
                php_code_contains_top_level_return(code)
            }
            MixedPhpIncludeSegment::Html(_) => false,
        })
}

fn php_code_contains_top_level_return(code: &str) -> bool {
    #[derive(Copy, Clone, Eq, PartialEq)]
    enum ScanState {
        Normal,
        SingleQuote,
        DoubleQuote,
        LineComment,
        BlockComment,
    }

    fn is_ident_byte(byte: u8) -> bool {
        matches!(byte, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_')
    }

    let bytes = code.as_bytes();
    let mut state = ScanState::Normal;
    let mut index = 0usize;
    let mut brace_depth = 0usize;

    while index < bytes.len() {
        match state {
            ScanState::Normal => {
                if let Some(next_index) = skip_php_heredoc(bytes, index) {
                    index = next_index;
                    continue;
                }
                match bytes[index] {
                    b'\'' => {
                        state = ScanState::SingleQuote;
                        index += 1;
                    }
                    b'"' => {
                        state = ScanState::DoubleQuote;
                        index += 1;
                    }
                    b'/' if bytes.get(index + 1) == Some(&b'/') => {
                        state = ScanState::LineComment;
                        index += 2;
                    }
                    b'/' if bytes.get(index + 1) == Some(&b'*') => {
                        state = ScanState::BlockComment;
                        index += 2;
                    }
                    b'#' => {
                        state = ScanState::LineComment;
                        index += 1;
                    }
                    b'{' => {
                        brace_depth += 1;
                        index += 1;
                    }
                    b'}' => {
                        brace_depth = brace_depth.saturating_sub(1);
                        index += 1;
                    }
                    b'r' if bytes[index..].starts_with(b"return") => {
                        let prev_ok = index == 0 || !is_ident_byte(bytes[index - 1]);
                        let next_index = index + b"return".len();
                        let next_ok =
                            next_index >= bytes.len() || !is_ident_byte(bytes[next_index]);
                        if brace_depth == 0 && prev_ok && next_ok {
                            return true;
                        }
                        index += 1;
                    }
                    _ => index += 1,
                }
            }
            ScanState::SingleQuote => {
                if bytes[index] == b'\\' {
                    index = (index + 2).min(bytes.len());
                } else if bytes[index] == b'\'' {
                    state = ScanState::Normal;
                    index += 1;
                } else {
                    index += 1;
                }
            }
            ScanState::DoubleQuote => {
                if bytes[index] == b'\\' {
                    index = (index + 2).min(bytes.len());
                } else if bytes[index] == b'"' {
                    state = ScanState::Normal;
                    index += 1;
                } else {
                    index += 1;
                }
            }
            ScanState::LineComment => {
                if bytes[index] == b'\n' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::BlockComment => {
                if bytes[index] == b'*' && bytes.get(index + 1) == Some(&b'/') {
                    state = ScanState::Normal;
                    index += 2;
                } else {
                    index += 1;
                }
            }
        }
    }

    false
}

fn rewrite_php_magic_constants(source: &str, source_path: &Path) -> String {
    if !source.contains("__DIR__") && !source.contains("__FILE__") {
        return source.to_string();
    }

    if !source.contains("<?") {
        return rewrite_php_magic_constants_in_code(source, source_path);
    }

    let mut out = String::new();
    for segment in split_mixed_php_include_source(source) {
        match segment {
            MixedPhpIncludeSegment::Html(html) => out.push_str(html),
            MixedPhpIncludeSegment::Echo {
                expr,
                has_close_tag,
            } => {
                out.push_str("<?=");
                out.push_str(&rewrite_php_magic_constants_in_code(expr, source_path));
                if has_close_tag {
                    out.push_str("?>");
                }
            }
            MixedPhpIncludeSegment::Code {
                code,
                has_close_tag,
            } => {
                out.push_str("<?php");
                out.push_str(&rewrite_php_magic_constants_in_code(code, source_path));
                if has_close_tag {
                    out.push_str("?>");
                }
            }
        }
    }
    out
}

fn rewrite_php_magic_constants_in_code(code: &str, source_path: &Path) -> String {
    #[derive(Copy, Clone, Eq, PartialEq)]
    enum ScanState {
        Normal,
        SingleQuote,
        DoubleQuote,
        LineComment,
        BlockComment,
    }

    fn is_ident_byte(byte: u8) -> bool {
        matches!(byte, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_')
    }

    fn php_single_quoted_literal(text: &str) -> String {
        let escaped = text.replace('\\', "\\\\").replace('\'', "\\'");
        format!("'{}'", escaped)
    }

    fn matches_magic_token(bytes: &[u8], index: usize, token: &[u8]) -> bool {
        if !bytes[index..].starts_with(token) {
            return false;
        }
        let prev_ok = index == 0 || !is_ident_byte(bytes[index - 1]);
        let next_index = index + token.len();
        let next_ok = next_index >= bytes.len() || !is_ident_byte(bytes[next_index]);
        prev_ok && next_ok
    }

    let dir_value = php_single_quoted_literal(
        &absolutize_path(source_path)
            .parent()
            .unwrap_or(Path::new("."))
            .to_string_lossy(),
    );
    let file_value = php_single_quoted_literal(&absolutize_path(source_path).to_string_lossy());

    let bytes = code.as_bytes();
    // Accumulate raw bytes so multibyte UTF-8 sequences survive verbatim; the
    // scanner only branches on ASCII delimiters, so byte-wise copying is safe.
    let mut out: Vec<u8> = Vec::with_capacity(code.len());
    let mut index = 0usize;
    let mut state = ScanState::Normal;

    while index < bytes.len() {
        match state {
            ScanState::Normal => {
                if let Some(next_index) = skip_php_heredoc(bytes, index) {
                    out.extend_from_slice(code[index..next_index].as_bytes());
                    index = next_index;
                    continue;
                }
                if matches_magic_token(bytes, index, b"__DIR__") {
                    out.extend_from_slice(dir_value.as_bytes());
                    index += "__DIR__".len();
                    continue;
                }
                if matches_magic_token(bytes, index, b"__FILE__") {
                    out.extend_from_slice(file_value.as_bytes());
                    index += "__FILE__".len();
                    continue;
                }
                if bytes[index] == b'\'' {
                    state = ScanState::SingleQuote;
                } else if bytes[index] == b'"' {
                    state = ScanState::DoubleQuote;
                } else if bytes[index] == b'#' {
                    state = ScanState::LineComment;
                } else if index + 1 < bytes.len()
                    && bytes[index] == b'/'
                    && bytes[index + 1] == b'/'
                {
                    state = ScanState::LineComment;
                } else if index + 1 < bytes.len()
                    && bytes[index] == b'/'
                    && bytes[index + 1] == b'*'
                {
                    state = ScanState::BlockComment;
                }
                out.push(bytes[index]);
                index += 1;
            }
            ScanState::SingleQuote => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    index += 1;
                    out.push(bytes[index]);
                } else if bytes[index] == b'\'' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::DoubleQuote => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    index += 1;
                    out.push(bytes[index]);
                } else if bytes[index] == b'"' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::LineComment => {
                out.push(bytes[index]);
                if bytes[index] == b'\n' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::BlockComment => {
                out.push(bytes[index]);
                if index + 1 < bytes.len() && bytes[index] == b'*' && bytes[index + 1] == b'/' {
                    index += 1;
                    out.push(bytes[index]);
                    state = ScanState::Normal;
                }
                index += 1;
            }
        }
    }

    String::from_utf8(out).unwrap_or_else(|_| code.to_string())
}

enum MixedPhpIncludeSegment<'a> {
    Html(&'a str),
    Code { code: &'a str, has_close_tag: bool },
    Echo { expr: &'a str, has_close_tag: bool },
}

fn php_code_block_needs_terminator(code: &str) -> bool {
    let trimmed = code.trim_end();
    let Some(last) = trimmed.chars().last() else {
        return false;
    };
    !matches!(last, ';' | '{' | '}' | ':')
}

fn split_mixed_php_include_source(source: &str) -> Vec<MixedPhpIncludeSegment<'_>> {
    let mut segments = Vec::new();
    let mut cursor = 0usize;

    while let Some(open_rel) = source[cursor..].find("<?") {
        let open = cursor + open_rel;
        if open > cursor {
            segments.push(MixedPhpIncludeSegment::Html(&source[cursor..open]));
        }

        let is_echo = source[open..].starts_with("<?=");
        let code_start = if is_echo {
            open + 3
        } else if source[open..].starts_with("<?php") {
            open + 5
        } else {
            open + 2
        };
        let close = find_php_include_close_tag(source, code_start).unwrap_or(source.len());
        let has_close_tag = close < source.len();
        let code = &source[code_start..close];
        if is_echo {
            segments.push(MixedPhpIncludeSegment::Echo {
                expr: code,
                has_close_tag,
            });
        } else {
            segments.push(MixedPhpIncludeSegment::Code {
                code,
                has_close_tag,
            });
        }
        cursor = if has_close_tag {
            (close + 2).min(source.len())
        } else {
            close
        };
    }

    if cursor < source.len() {
        segments.push(MixedPhpIncludeSegment::Html(&source[cursor..]));
    }

    segments
}

fn skip_php_heredoc(bytes: &[u8], start: usize) -> Option<usize> {
    if start + 3 >= bytes.len()
        || bytes[start] != b'<'
        || bytes[start + 1] != b'<'
        || bytes[start + 2] != b'<'
    {
        return None;
    }

    let mut index = start + 3;
    while index < bytes.len() && matches!(bytes[index], b' ' | b'\t') {
        index += 1;
    }

    let quote = match bytes.get(index).copied() {
        Some(b'\'') | Some(b'"') => {
            let q = bytes[index];
            index += 1;
            Some(q)
        }
        _ => None,
    };

    let tag_start = index;
    if !matches!(
        bytes.get(index).copied(),
        Some(b'A'..=b'Z' | b'a'..=b'z' | b'_')
    ) {
        return None;
    }
    index += 1;
    while index < bytes.len()
        && matches!(bytes[index], b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_')
    {
        index += 1;
    }
    let tag = &bytes[tag_start..index];

    if let Some(q) = quote {
        if bytes.get(index).copied() != Some(q) {
            return None;
        }
        index += 1;
    }

    while index < bytes.len() {
        match bytes[index] {
            b'\n' => {
                index += 1;
                break;
            }
            b'\r' => {
                index += 1;
                if bytes.get(index).copied() == Some(b'\n') {
                    index += 1;
                }
                break;
            }
            _ => index += 1,
        }
    }

    while index < bytes.len() {
        let line_start = index;
        while index < bytes.len() && matches!(bytes[index], b' ' | b'\t') {
            index += 1;
        }

        if bytes[index..].starts_with(tag) {
            let mut after = index + tag.len();
            if after == bytes.len() || matches!(bytes[after], b';' | b'\n' | b'\r') {
                if bytes.get(after).copied() == Some(b';') {
                    after += 1;
                }
                while after < bytes.len() {
                    match bytes[after] {
                        b'\n' => {
                            after += 1;
                            break;
                        }
                        b'\r' => {
                            after += 1;
                            if bytes.get(after).copied() == Some(b'\n') {
                                after += 1;
                            }
                            break;
                        }
                        _ => after += 1,
                    }
                }
                return Some(after);
            }
        }

        index = line_start;
        while index < bytes.len() {
            match bytes[index] {
                b'\n' => {
                    index += 1;
                    break;
                }
                b'\r' => {
                    index += 1;
                    if bytes.get(index).copied() == Some(b'\n') {
                        index += 1;
                    }
                    break;
                }
                _ => index += 1,
            }
        }
    }

    Some(bytes.len())
}

fn find_php_include_close_tag(source: &str, start: usize) -> Option<usize> {
    #[derive(Copy, Clone, Eq, PartialEq)]
    enum ScanState {
        Normal,
        SingleQuote,
        DoubleQuote,
        LineComment,
        BlockComment,
    }

    let bytes = source.as_bytes();
    let mut index = start;
    let mut state = ScanState::Normal;

    while index + 1 < bytes.len() {
        match state {
            ScanState::Normal => {
                if let Some(next_index) = skip_php_heredoc(bytes, index) {
                    index = next_index;
                    continue;
                }
                if bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    return Some(index);
                }
                if bytes[index] == b'\'' {
                    state = ScanState::SingleQuote;
                } else if bytes[index] == b'"' {
                    state = ScanState::DoubleQuote;
                } else if bytes[index] == b'#' {
                    state = ScanState::LineComment;
                } else if bytes[index] == b'/' && bytes[index + 1] == b'/' {
                    state = ScanState::LineComment;
                    index += 1;
                } else if bytes[index] == b'/' && bytes[index + 1] == b'*' {
                    state = ScanState::BlockComment;
                    index += 1;
                }
            }
            ScanState::SingleQuote => {
                if bytes[index] == b'\\' {
                    index += 1;
                } else if bytes[index] == b'\'' {
                    state = ScanState::Normal;
                }
            }
            ScanState::DoubleQuote => {
                if bytes[index] == b'\\' {
                    index += 1;
                } else if bytes[index] == b'"' {
                    state = ScanState::Normal;
                }
            }
            ScanState::LineComment => {
                if bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    return Some(index);
                }
                if bytes[index] == b'\n' {
                    state = ScanState::Normal;
                }
            }
            ScanState::BlockComment => {
                if bytes[index] == b'*' && bytes[index + 1] == b'/' {
                    state = ScanState::Normal;
                    index += 1;
                }
            }
        }
        index += 1;
    }

    None
}

fn normalize_php_alternative_control_syntax(line: &str) -> String {
    #[derive(Copy, Clone, Eq, PartialEq)]
    enum TagMode {
        Bare,
        OpenOnly,
        CloseOnly,
        Wrapped,
    }

    fn split_line_comment<'a>(line: &'a str) -> (&'a str, &'a str) {
        if let Some((head, tail)) = line.split_once("//") {
            return (head.trim_end(), tail);
        }
        if let Some((head, tail)) = line.split_once('#') {
            return (head.trim_end(), tail);
        }
        (line.trim_end(), "")
    }

    fn starts_control_header(prefix: &str, keyword: &str) -> bool {
        prefix.starts_with(&format!("{keyword} ")) || prefix.starts_with(&format!("{keyword}("))
    }

    let indent_len = line.len() - line.trim_start().len();
    let indent = &line[..indent_len];
    let trimmed = line.trim();
    let (inner, tag_mode, close_suffix) = if let Some(rest) = trimmed.strip_prefix("<?php") {
        let rest = rest.trim();
        if let Some(close_index) = rest.find("?>") {
            let inner = rest[..close_index].trim();
            let close_suffix = &rest[close_index + 2..];
            (inner, TagMode::Wrapped, close_suffix)
        } else {
            (rest, TagMode::OpenOnly, "")
        }
    } else if let Some(close_index) = trimmed.find("?>") {
        let inner = trimmed[..close_index].trim();
        let close_suffix = &trimmed[close_index + 2..];
        (inner, TagMode::CloseOnly, close_suffix)
    } else {
        (trimmed, TagMode::Bare, "")
    };
    let (code, comment) = split_line_comment(inner);
    let code = code.trim();
    let comment_suffix = if comment.is_empty() {
        String::new()
    } else if inner.contains("//") {
        format!(" //{}", comment)
    } else {
        format!(" #{}", comment)
    };

    for end_keyword in ["endif", "endforeach", "endfor", "endwhile", "endswitch"] {
        let end_prefix = format!("{end_keyword};");
        if let Some(rest) = code.strip_prefix(&end_prefix) {
            let rest = rest.trim_start();
            if !rest.is_empty() {
                let normalized_rest = normalize_php_alternative_control_syntax(rest);
                let combined = format!("}} {}", normalized_rest.trim_start());
                return match tag_mode {
                    TagMode::Wrapped => format!("{}<?php {} ?>{}", indent, combined, close_suffix),
                    TagMode::OpenOnly => format!("{}<?php {}{}", indent, combined, comment_suffix),
                    TagMode::CloseOnly => format!("{}{} ?>{}", indent, combined, close_suffix),
                    TagMode::Bare => format!("{}{}{}", indent, combined, comment_suffix),
                };
            }
        }
    }

    if let Some(prefix) = code.strip_suffix(':').map(str::trim_end) {
        if starts_control_header(prefix, "if")
            || starts_control_header(prefix, "foreach")
            || starts_control_header(prefix, "for")
            || starts_control_header(prefix, "while")
            || starts_control_header(prefix, "switch")
        {
            return match tag_mode {
                TagMode::Wrapped => format!("{}<?php {} {{ ?>{}", indent, prefix, close_suffix),
                TagMode::OpenOnly => format!("{}<?php {} {{{}", indent, prefix, comment_suffix),
                TagMode::CloseOnly => format!("{}{} {{ ?>{}", indent, prefix, close_suffix),
                TagMode::Bare => format!("{}{} {{{}", indent, prefix, comment_suffix),
            };
        }
        if prefix.starts_with("elseif ") {
            return match tag_mode {
                TagMode::Wrapped => format!("{}<?php }} {} {{ ?>{}", indent, prefix, close_suffix),
                TagMode::OpenOnly => format!("{}<?php }} {} {{{}", indent, prefix, comment_suffix),
                TagMode::CloseOnly => format!("{}}} {} {{ ?>{}", indent, prefix, close_suffix),
                TagMode::Bare => format!("{}}} {} {{{}", indent, prefix, comment_suffix),
            };
        }
        if prefix == "else" {
            return match tag_mode {
                TagMode::Wrapped => format!("{}<?php }} else {{ ?>{}", indent, close_suffix),
                TagMode::OpenOnly => format!("{}<?php }} else {{{}", indent, comment_suffix),
                TagMode::CloseOnly => format!("{}}} else {{ ?>{}", indent, close_suffix),
                TagMode::Bare => format!("{}}} else {{{}", indent, comment_suffix),
            };
        }
    }

    let end_keyword = code.strip_suffix(';').map(str::trim_end).unwrap_or(code);
    if matches!(
        end_keyword,
        "endif" | "endforeach" | "endfor" | "endwhile" | "endswitch"
    ) {
        return match tag_mode {
            TagMode::Wrapped => format!("{}<?php }} ?>{}", indent, close_suffix),
            TagMode::OpenOnly => format!("{}<?php }}{}", indent, comment_suffix),
            TagMode::CloseOnly => format!("{}}} ?>{}", indent, close_suffix),
            TagMode::Bare => format!("{}}}{}", indent, comment_suffix),
        };
    }

    line.to_string()
}

fn starts_multiline_php_alternative_control_header(line: &str) -> bool {
    fn starts_control_header(prefix: &str, keyword: &str) -> bool {
        prefix.starts_with(&format!("{keyword} ")) || prefix.starts_with(&format!("{keyword}("))
    }

    let trimmed = line.trim();
    let body = trimmed
        .strip_prefix("<?php")
        .map(str::trim_start)
        .unwrap_or(trimmed);

    let starts_for = starts_control_header(body, "for");
    let starts_control = starts_control_header(body, "if")
        || starts_control_header(body, "foreach")
        || starts_for
        || starts_control_header(body, "while")
        || starts_control_header(body, "switch");

    starts_control
        && !body.contains('{')
        && !body.contains(':')
        && !body.contains("?>")
        && (starts_for || !body.contains(';'))
}

fn should_stop_multiline_php_header_scan(line: &str) -> bool {
    let trimmed = line.trim_start();
    trimmed.starts_with("case ")
        || trimmed.starts_with("default:")
        || line.contains('{')
        || line.contains('}')
}

fn normalize_multiline_php_alternative_control_syntax(block: &str) -> Option<String> {
    let lines: Vec<&str> = block.lines().collect();
    if lines.len() < 2 {
        return None;
    }

    if lines.iter().any(|line| {
        let trimmed = line.trim_start();
        trimmed.starts_with("case ")
            || trimmed.starts_with("default:")
            || line.contains('{')
            || line.contains('}')
    }) {
        return None;
    }

    let first_trimmed = lines.first()?.trim();
    let first_body = first_trimmed
        .strip_prefix("<?php")
        .map(str::trim_start)
        .unwrap_or(first_trimmed);
    let last_trimmed = lines.last()?.trim();
    let (last_body, close_suffix) = if let Some(close_index) = last_trimmed.find("?>") {
        (
            last_trimmed[..close_index].trim_end(),
            &last_trimmed[close_index + 2..],
        )
    } else {
        (last_trimmed, "")
    };

    let mut combined = String::from(first_body);
    for middle in lines.iter().skip(1).take(lines.len().saturating_sub(2)) {
        combined.push('\n');
        combined.push_str(middle.trim());
    }
    combined.push('\n');
    combined.push_str(last_body);

    let prefix = combined.trim_end().strip_suffix(':')?.trim_end();
    if !(prefix.starts_with("if ")
        || prefix.starts_with("foreach ")
        || prefix.starts_with("for ")
        || prefix.starts_with("while ")
        || prefix.starts_with("switch "))
    {
        return None;
    }

    let last_line = *lines.last()?;
    let last_indent_len = last_line.len() - last_line.trim_start().len();
    let last_indent = &last_line[..last_indent_len];
    let last_prefix = last_body.trim_end().strip_suffix(':')?.trim_end();

    let mut normalized_lines = lines
        .iter()
        .map(|line| line.to_string())
        .collect::<Vec<_>>();
    let mut normalized_last = format!("{}{} {{", last_indent, last_prefix);
    if last_trimmed.contains("?>") {
        normalized_last.push_str(" ?>");
        normalized_last.push_str(close_suffix);
    }
    *normalized_lines.last_mut()? = normalized_last;
    Some(normalized_lines.join("\n"))
}

fn normalize_php_source_for_parser(source: &str) -> String {
    let source = rewrite_php_execution_operator(source);
    let mut normalized_lines = Vec::new();
    let mut lines = source.lines().peekable();

    while let Some(line) = lines.next() {
        if starts_multiline_php_alternative_control_header(line) {
            let mut header_lines = vec![line.to_string()];
            let mut normalized_block = None;
            let mut scanned_lines = 0usize;

            while let Some(next_line) = lines.peek().copied() {
                if should_stop_multiline_php_header_scan(next_line) || scanned_lines >= 12 {
                    break;
                }
                header_lines.push(next_line.to_string());
                lines.next();
                scanned_lines += 1;
                let candidate = header_lines.join("\n");
                if let Some(normalized) =
                    normalize_multiline_php_alternative_control_syntax(&candidate)
                {
                    normalized_block = Some(normalized);
                    break;
                }
            }

            normalized_lines.extend(
                normalized_block
                    .unwrap_or_else(|| header_lines.join("\n"))
                    .lines()
                    .map(str::to_string),
            );
            continue;
        }

        normalized_lines.push(normalize_php_alternative_control_syntax(line));
    }

    let mut normalized = normalized_lines.join("\n");
    if source.ends_with('\n') {
        normalized.push('\n');
    }

    let mixed_normalized = match vybe_bytecode::registry::hooks("php").normalize_source {
        Some(f) => f(&normalized),
        None => normalized,
    };
    let mut normalized = mixed_normalized
        .lines()
        .map(normalize_php_alternative_control_syntax)
        .collect::<Vec<_>>()
        .join("\n");
    if mixed_normalized.ends_with('\n') {
        normalized.push('\n');
    }
    normalized
}

fn rewrite_php_execution_operator(source: &str) -> String {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum State {
        Normal,
        SingleQuoted,
        DoubleQuoted,
        LineComment,
        BlockComment,
    }

    let bytes = source.as_bytes();
    // Accumulate raw bytes so multibyte UTF-8 sequences survive verbatim; the
    // scanner only branches on ASCII delimiters, so byte-wise copying is safe.
    let mut out: Vec<u8> = Vec::with_capacity(source.len());
    let mut index = 0usize;
    let mut state = State::Normal;

    while index < bytes.len() {
        match state {
            State::Normal => {
                if bytes[index] == b'`' {
                    index += 1;
                    let mut inner: Vec<u8> = Vec::new();

                    while index < bytes.len() {
                        let byte = bytes[index];
                        if byte == b'\\' && index + 1 < bytes.len() {
                            let next = bytes[index + 1];
                            if next == b'"' {
                                inner.push(b'\\');
                            }
                            inner.push(next);
                            index += 2;
                            continue;
                        }
                        if byte == b'`' {
                            index += 1;
                            break;
                        }
                        if byte == b'"' {
                            inner.push(b'\\');
                        }
                        inner.push(byte);
                        index += 1;
                    }

                    out.extend_from_slice(b"shell_exec(\"");
                    out.extend_from_slice(&inner);
                    out.extend_from_slice(b"\")");
                    continue;
                }

                if bytes[index] == b'\'' {
                    state = State::SingleQuoted;
                } else if bytes[index] == b'"' {
                    state = State::DoubleQuoted;
                } else if bytes[index] == b'/'
                    && index + 1 < bytes.len()
                    && bytes[index + 1] == b'/'
                {
                    state = State::LineComment;
                } else if bytes[index] == b'#' {
                    state = State::LineComment;
                } else if bytes[index] == b'/'
                    && index + 1 < bytes.len()
                    && bytes[index + 1] == b'*'
                {
                    state = State::BlockComment;
                }

                out.push(bytes[index]);
                index += 1;
            }
            State::SingleQuoted => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1]);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'\'' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::DoubleQuoted => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1]);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'"' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::LineComment => {
                out.push(bytes[index]);
                if bytes[index] == b'\n' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::BlockComment => {
                out.push(bytes[index]);
                if bytes[index] == b'*' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    out.push(b'/');
                    index += 2;
                    state = State::Normal;
                    continue;
                }
                index += 1;
            }
        }
    }

    String::from_utf8(out).unwrap_or_else(|_| source.to_string())
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
        // Host Component Model namespaces (`wasi:*`, `wasm:*`, `vybe:*`,
        // `node:*`) are not source files — the compiler binds them at
        // compile time from `module.imports` directly via
        // `host_import_bindings`. Skip filesystem resolution so we
        // don't print spurious "no such file" warnings.
        if path_str.starts_with("wasi:")
            || path_str.starts_with("wasm:")
            || path_str.starts_with("vybe:")
            || path_str.starts_with("node:")
            // Flutter is a resolver-provided pseudo-package (its `flutter.*`
            // widgets are registered in the namespace tree + the adapter
            // runtime), not a source file on disk.
            || path_str.starts_with("package:flutter/")
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

#[cfg(test)]
mod tests {
    /// Register the frontends these tests need. `vybe_compiler` links no
    /// language crate in production — languages register THEMSELVES through
    /// `vybe_bytecode::registry` — so this crate's test binary starts with an
    /// empty registry. This is the same `register()` entry point vybex and a
    /// dylib host would call.
    fn register_test_languages() {
        vybe_language_php::register();
        vybe_language_cobol::register();
    }

    use super::*;

    fn run_php_bundle_prints(bundle: &Bundle) -> Vec<String> {
        use std::sync::{Arc, Mutex};
        use vybe_bytecode::{HostContext, VM, Value};

        let compiled = bundle.compile_full().expect("compiled bundle");
        let mut vm = VM::new();
        let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
        let captured = output.clone();

        crate::compiler::platforms::register_platforms_all(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
            "log",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let msg = match args.len() {
                    0 => String::new(),
                    1 => format!("{}", args[0]),
                    2 => format!("{}", args[1]),
                    _ => format!("{}", args[2]),
                };
                captured.lock().unwrap().push(msg);
                Value::Null
            }),
        );
        vm.run(compiled.chunks).expect("run bundle");

        output.lock().unwrap().clone()
    }

    #[test]
    fn php_bundle_expands_require_once_with_dir_alias() {
        let temp_root =
            std::env::temp_dir().join(format!("vybex_php_bundle_{}", uuid::Uuid::new_v4()));
        let public_dir = temp_root.join("public");
        std::fs::create_dir_all(&public_dir).expect("create temp dirs");

        let lib_path = temp_root.join("lib.php");
        std::fs::write(&lib_path, "<?php\nfunction shared_helper() { return 1; }\n")
            .expect("write lib");

        let entry_path = public_dir.join("entry.php");
        let entry_src = "<?php\n$basedir = dirname(__DIR__);\nrequire_once $basedir . '/lib.php';\necho shared_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "shared_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_rewrites_magic_dir_and_file_constants_for_runtime() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_magic_consts_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\necho __DIR__;\necho __FILE__;\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let outputs = run_php_bundle_prints(&bundle);
        assert_eq!(
            outputs,
            vec![
                temp_root.to_string_lossy().into_owned(),
                entry_path.to_string_lossy().into_owned(),
            ]
        );

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn cobol_bundle_rewrites_assign_to_paths_relative_to_source() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_cobol_bundle_assign_to_{}",
            uuid::Uuid::new_v4()
        ));
        let data_dir = temp_root.join("fixtures");
        std::fs::create_dir_all(&data_dir).expect("create temp dirs");

        let entry_path = data_dir.join("entry.cbl");
        let entry_src = r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT CUSTOMER-FILE ASSIGN TO "customers.dat".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DUMMY PIC X.
PROCEDURE DIVISION.
    DISPLAY "ok".
    STOP RUN.
"#;
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("cobol").expect("cobol language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let combined = rewrite_cobol_assign_paths(&bundle.sources[0].code, &bundle.sources[0].path);
        assert!(combined.contains(&format!(
            "ASSIGN TO \"{}\"",
            data_dir.join("customers.dat").to_string_lossy()
        )));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_expands_require_once_with_defined_constants() {
        let temp_root =
            std::env::temp_dir().join(format!("vybex_php_bundle_define_{}", uuid::Uuid::new_v4()));
        let includes_dir = temp_root.join("includes");
        std::fs::create_dir_all(&includes_dir).expect("create temp dirs");

        let lib_path = includes_dir.join("shared.php");
        std::fs::write(
            &lib_path,
            "<?php\nfunction shared_constant_helper() { return 7; }\n",
        )
        .expect("write lib");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\ndefine('WPINC', 'includes');\nrequire_once ABSPATH . WPINC . '/shared.php';\necho shared_constant_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "shared_constant_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_resolves_nested_bare_relative_include_from_entry_dir() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_entry_relative_{}",
            uuid::Uuid::new_v4()
        ));
        let classes_dir = temp_root.join("classes");
        let views_dir = temp_root.join("views/users");
        std::fs::create_dir_all(&classes_dir).expect("create classes dir");
        std::fs::create_dir_all(&views_dir).expect("create views dir");

        let entry_path = temp_root.join("index.php");
        let controller_path = classes_dir.join("UserController.php");
        let view_path = views_dir.join("list.php");

        std::fs::write(
            &entry_path,
            "<?php\nrequire_once 'classes/UserController.php';\n",
        )
        .expect("write entry");
        std::fs::write(&controller_path, "<?php\nfunction render_users() {\n    include 'views/users/list.php';\n    return users_view_value();\n}\n")
            .expect("write controller");
        std::fs::write(
            &view_path,
            "<?php\nfunction users_view_value() { return 42; }\n",
        )
        .expect("write view");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: std::fs::read_to_string(&entry_path).expect("read entry"),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let expanded =
            expand_php_bundle_sources_with_map(&bundle.sources).expect("expand php bundle");
        let expanded_code = expanded.into_code();
        assert!(expanded_code.contains("function render_users()"));
        assert!(expanded_code.contains("function users_view_value()"));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_nested_mixed_html_include_inside_class_method() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_nested_mixed_include_{}",
            uuid::Uuid::new_v4()
        ));
        let classes_dir = temp_root.join("classes");
        let views_dir = temp_root.join("views/issues");
        std::fs::create_dir_all(&classes_dir).expect("create classes dir");
        std::fs::create_dir_all(&views_dir).expect("create views dir");

        let entry_path = temp_root.join("index.php");
        let controller_path = classes_dir.join("IssueController.php");
        let view_path = views_dir.join("edit.php");

        let entry_src = "<?php\nrequire_once 'classes/IssueController.php';\n";
        let controller_src = "<?php\nclass IssueController {\n    public function edit($issue) {\n        include 'views/issues/edit.php';\n    }\n}\n";
        let view_src = "<?php $pageTitle = 'Edit'; ?>\n<div><?= htmlspecialchars($issue['SUMMARY']) ?></div>\n";

        std::fs::write(&entry_path, entry_src).expect("write entry");
        std::fs::write(&controller_path, controller_src).expect("write controller");
        std::fs::write(&view_path, view_src).expect("write view");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::ClassDecl { name, .. } if name == "IssueController")));

        let expanded =
            expand_php_bundle_sources_with_map(&bundle.sources).expect("expand php bundle");
        let expanded_code = expanded.into_code();
        assert!(expanded_code.contains("include 'views/issues/edit.php';"));
        assert!(!expanded_code.contains("<div>"));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_nested_pure_php_require_once_inside_class_method() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_nested_require_once_{}",
            uuid::Uuid::new_v4()
        ));
        let classes_dir = temp_root.join("classes");
        let vendor_dir = temp_root.join("vendor");
        std::fs::create_dir_all(&classes_dir).expect("create classes dir");
        std::fs::create_dir_all(&vendor_dir).expect("create vendor dir");

        let entry_path = temp_root.join("index.php");
        let controller_path = classes_dir.join("IssueController.php");
        let include_path = vendor_dir.join("Loader.php");

        let entry_src = "<?php\nrequire_once 'classes/IssueController.php';\n";
        let controller_src = "<?php\nclass IssueController {\n    public static function load() {\n        require_once __DIR__ . '/../vendor/Loader.php';\n    }\n}\n";
        let include_src = "<?php\nclass Loader {}\n";

        std::fs::write(&entry_path, entry_src).expect("write entry");
        std::fs::write(&controller_path, controller_src).expect("write controller");
        std::fs::write(&include_path, include_src).expect("write include");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let expanded =
            expand_php_bundle_sources_with_map(&bundle.sources).expect("expand php bundle");
        let expanded_code = expanded.into_code();
        assert!(expanded_code.contains("require_once __DIR__ . '/../vendor/Loader.php';"));
        assert!(!expanded_code.contains("class Loader {}"));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_dynamic_variable_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_dynamic_include_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction load_dynamic($file) {\n    require_once $file;\n}\nfunction still_present() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "load_dynamic")));
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "still_present")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_top_level_returning_require_once() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_returning_require_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let autoload_path = temp_root.join("autoload.php");
        std::fs::write(
            &autoload_path,
            "<?php\nfunction loader_value() { return 7; }\nreturn loader_value();\n",
        )
        .expect("write autoload");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nrequire_once 'autoload.php';\necho 'done';\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let expanded =
            expand_php_bundle_sources_with_map(&bundle.sources).expect("expand php bundle");
        let expanded_code = expanded.into_code();
        assert!(expanded_code.contains("require_once 'autoload.php';"));
        assert!(!expanded_code.contains("return loader_value();"));
        assert!(expanded_code.contains("echo 'done';"));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_parses_define_with_trailing_inline_comment() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_define_comment_{}",
            uuid::Uuid::new_v4()
        ));
        let includes_dir = temp_root.join("wp-content");
        std::fs::create_dir_all(&includes_dir).expect("create temp dirs");

        let lib_path = includes_dir.join("db-error.php");
        std::fs::write(
            &lib_path,
            "<?php\nfunction db_error_helper() { return 9; }\n",
        )
        .expect("write lib");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\ndefine('WP_CONTENT_DIR', ABSPATH . 'wp-content'); // trailing comment\nrequire_once WP_CONTENT_DIR . '/db-error.php';\necho db_error_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "db_error_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_unknown_constant_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_unknown_const_include_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nif (file_exists(WP_CONTENT_DIR . '/db-error.php')) {\n    require_once WP_CONTENT_DIR . '/db-error.php';\n}\nfunction after_optional_include() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_optional_include")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_missing_include_when_guarded_by_file_exists() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_optional_missing_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\nif ( ! file_exists( ABSPATH . '.maintenance' ) ) {\n    return;\n}\nrequire ABSPATH . '.maintenance';\nfunction after_optional_file() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_optional_file")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_function_local_alias_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_local_alias_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction load_runtime($engine) {\n    $file = __DIR__ . '/' . $engine . '.php';\n    require_once $file;\n}\nfunction after_runtime_loader() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "load_runtime")));
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_runtime_loader")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_mixed_static_dynamic_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_mixed_dynamic_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\nfunction load_runtime($name) {\n    require_once ABSPATH . 'lib/' . $name . '.php';\n}\nfunction after_mixed_runtime_loader() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "load_runtime")));
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_mixed_runtime_loader")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_interpolated_double_quoted_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_interpolated_include_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('WP_LANG_DIR', __DIR__ . '/languages');\n$locale = get_locale();\n$locale_file = WP_LANG_DIR . \"/$locale.php\";\nif ( is_readable( $locale_file ) ) {\n    require $locale_file;\n}\nfunction after_locale_include() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_locale_include")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_does_not_treat_includes_prefixed_function_calls_as_include_statements() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_include_prefix_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction includes_url($path) { return $path; }\n$css = includes_url('style.css');\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_helper_call_include_paths() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_helper_call_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction render_template($base, $file) {\n    require trailingslashit($base) . $file;\n}\nfunction after_helper_include() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_helper_include")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_include_after_missing_positive_defined_guard() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_defined_guard_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nif ( defined('SUNRISE') ) {\n    include_once __DIR__ . '/sunrise.php';\n}\nfunction after_defined_guard() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_defined_guard")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_missing_plugin_include_when_guarded_by_is_plugin_active() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_plugin_guard_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('WP_PLUGIN_DIR', __DIR__ . '/wp-content/plugins');\nif ( is_plugin_active( 'press-this/press-this-plugin.php' ) ) {\n    include WP_PLUGIN_DIR . '/press-this/class-wp-press-this-plugin.php';\n}\nfunction after_plugin_guard() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_plugin_guard")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_alternative_template_if_syntax() {
        let temp_root =
            std::env::temp_dir().join(format!("vybex_php_bundle_alt_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src =
            "<?php\n$value = true;\n?>\n<?php if ($value): ?>\n<div>ok</div>\n<?php endif; ?>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_bare_alternative_while_syntax() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_while_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src =
            "<?php\n$i = 0;\nwhile ($i < 1) : ?>\n<div>ok</div>\n<?php\n$i++;\nendwhile;\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_bare_alternative_foreach_syntax() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_foreach_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\n$items = ['mp4'];\nforeach ($items as $type) : ?>\n<span><?php echo $type; ?></span>\n<?php\nendforeach;\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_wrapped_alternative_foreach_with_template_suffix() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_foreach_suffix_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nforeach ( ['autoplay'] as $attr ) :\n?>\n<# <?php echo $attr; ?> #>\n<?php endforeach; ?>#>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_multiline_alternative_foreach_header() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_foreach_multiline_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nforeach ( array(\n    'artist' => 'Artist',\n    'album' => 'Album',\n) as $key => $label ) :\n?>\n<span><?php echo $label; ?></span>\n<?php endforeach; ?>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_prepares_mixed_template_with_inner_function_and_template_loops() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_keynote_template_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = r#"<?php
$notesTree = [
    ['notes' => [['title' => 'One'], ['title' => 'Two']]],
];
?>
<div id="tree">
    <?php
    function renderTree($notes) {
        foreach ($notes as $note) {
            echo '<div class="tree-node">' . htmlspecialchars($note['title']) . '</div>';
        }
    }
    foreach ($notesTree as $section) {
        if (!empty($section['notes'])) {
            renderTree($section['notes']);
        }
    }
    ?>
</div>
<div id="sections">
    <?php foreach($notesTree as $section): ?>
        <span><?php echo count($section['notes']); ?></span>
    <?php endforeach; ?>
</div>
"#;
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_bare_alternative_endif_with_comment() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_endif_comment_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src =
            "<?php\nif (true) :\n?>\n<div>ok</div>\n<?php\nendif; // keep parser aligned\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_strips_included_close_tag_before_inline_html() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_close_tag_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let header_path = temp_root.join("header.php");
        std::fs::write(&header_path, "<?php echo 'head'; ?>\n<nav>nav</nav>\n")
            .expect("write header");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ninclude 'header.php';\n?>\n<div>body</div>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        let echoed_text: String = module
            .body
            .iter()
            .filter_map(|stmt| {
                if let StmtKind::Echo(exprs) = &stmt.kind {
                    if exprs.len() == 1 {
                        if let ExprKind::Lit(crate::ast::Literal::Str(text)) = &exprs[0].kind {
                            return Some(text.clone());
                        }
                    }
                }
                None
            })
            .collect();
        assert!(
            !echoed_text.contains("?>"),
            "bundled inline HTML should not contain a literal close tag: {echoed_text}"
        );

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_ignores_boundary_whitespace_from_included_code_file() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_boundary_ws_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let helper_path = temp_root.join("helper.php");
        std::fs::write(
            &helper_path,
            "\n<?php\nfunction helper_value() { return 1; }\n",
        )
        .expect("write helper");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ninclude 'helper.php';\nheader('Location: /next.php');\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        let echoed_text: String = module
            .body
            .iter()
            .filter_map(|stmt| {
                if let StmtKind::Echo(exprs) = &stmt.kind {
                    if exprs.len() == 1 {
                        if let ExprKind::Lit(crate::ast::Literal::Str(text)) = &exprs[0].kind {
                            return Some(text.clone());
                        }
                    }
                }
                None
            })
            .collect();
        assert!(
            echoed_text.is_empty(),
            "boundary whitespace from included code file should not become output: {echoed_text:?}"
        );

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_preserves_heredoc_xml_close_tags() {
        let temp_root =
            std::env::temp_dir().join(format!("vybex_php_bundle_heredoc_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let include_path = temp_root.join("ixr.php");
        let include_src = "<?php\nclass IXR_Request {\n    public $xml;\n    public function __construct() {\n        $this->xml = <<<EOD\n<?xml version=\"1.0\"?>\n<methodCall>\nEOD;\n    }\n}\n";
        std::fs::write(&include_path, include_src).expect("write include");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nrequire_once __DIR__ . '/ixr.php';\nnew IXR_Request();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let expanded = expand_php_bundle_sources(&bundle.sources).expect("expand php bundle");
        assert!(expanded.contains("<?xml version=\"1.0\"?>"));
        assert!(!expanded.contains("<?xml version=\"1.0\";?>"));
        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_does_not_rewrite_braced_switch_case_labels() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_switch_case_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction classify($mode) {\n    switch ($mode) {\n        case 0:\n            return 'zero';\n        default:\n            return 'other';\n    }\n}\necho classify(0);\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_prepares_wordpress_kses_file() {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let kses_path = manifest_dir
            .join("..")
            .join("..")
            .join("examples")
            .join("webroot")
            .join("wordpress")
            .join("wp-includes")
            .join("kses.php");
        let entry_src = std::fs::read_to_string(&kses_path).expect("read kses.php");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "kses".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: kses_path,
                code: entry_src,
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");
    }

    #[test]
    fn php_bundle_prepares_wordpress_index_entry() {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let index_path = manifest_dir
            .join("..")
            .join("..")
            .join("examples")
            .join("webroot")
            .join("wordpress")
            .join("index.php");
        let entry_src = std::fs::read_to_string(&index_path).expect("read index.php");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "wordpress-index".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: index_path,
                code: entry_src,
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");
    }

    #[test]
    fn php_bundle_prepares_wordpress_widgets_file() {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let widgets_path = manifest_dir
            .join("..")
            .join("..")
            .join("examples")
            .join("webroot")
            .join("wordpress")
            .join("wp-includes")
            .join("widgets.php");
        let source = std::fs::read_to_string(&widgets_path).expect("read widgets.php");
        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "widgets".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: widgets_path,
                code: source,
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");
    }

    #[test]
    fn php_bundle_prepares_wordpress_media_template_file() {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let media_template_path = manifest_dir
            .join("..")
            .join("..")
            .join("examples")
            .join("webroot")
            .join("wordpress")
            .join("wp-includes")
            .join("media-template.php");
        let source =
            std::fs::read_to_string(&media_template_path).expect("read media-template.php");
        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "media-template".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: media_template_path,
                code: source,
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");
    }

    #[test]
    fn php_bundle_rewrites_backtick_execution_operator() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_backtick_exec_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\n$commandline = 'printf ok';\n$result = `$commandline`;\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_prepares_doc_comment_with_php_tag_example() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_comment_tags_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\n/**\n * Example:\n * <main><p><?php echo \"Hello\"; ?></p></main>\n */\nfunction demo() {\n    $can_use_cached = ! wp_is_development_mode( 'theme' );\n    return $can_use_cached;\n}\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_prepares_readonly_named_function() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_readonly_fn_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction readonly( $readonly_value, $current = true, $display = true ) {\n    return $readonly_value;\n}\nreadonly( true );\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_chained_endif_if_same_line() {
        let temp_root = std::env::temp_dir().join(format!(
            "vybex_php_bundle_alt_chain_{}",
            uuid::Uuid::new_v4()
        ));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\n$first = true;\n$second = true;\nif ( $first ) : ?>\n<div>first</div>\n<?php endif; if ( $second ) : ?>\n<div>second</div>\n<?php endif; ?>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        register_test_languages();
        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile {
                path: entry_path.clone(),
                code: entry_src.to_string(),
            }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }
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

/// Extract host constant values from module records.
///
/// Returns `module → (export_name → Value)` for all `ExportEntry::Value`
/// entries, following Indirect chains. Used by the compiler to distinguish
/// constant-valued imports (e.g. `ecma:math::PI`) from callable function
/// imports at compile time, so constants can be inlined rather than routed
/// through `CALL_IMPORT`.
pub fn flatten_module_value_exports(
    modules: &HashMap<String, ModuleRecord>,
) -> HashMap<String, HashMap<String, vybe_bytecode::Value>> {
    let mut out: HashMap<String, HashMap<String, vybe_bytecode::Value>> = HashMap::new();
    for (specifier, record) in modules {
        for (name, _entry) in &record.exports {
            // Follow Indirect chains
            let terminal = {
                let mut visited: Vec<(String, String)> = Vec::new();
                let mut cur_mod = specifier.as_str();
                let mut cur_name = name.as_str();
                let mut result = None;
                loop {
                    if visited.contains(&(cur_mod.to_string(), cur_name.to_string())) {
                        break;
                    }
                    visited.push((cur_mod.to_string(), cur_name.to_string()));
                    match modules.get(cur_mod).and_then(|r| r.exports.get(cur_name)) {
                        Some(ExportEntry::Value(v)) => {
                            result = Some((cur_mod.to_string(), cur_name.to_string(), v.clone()));
                            break;
                        }
                        Some(ExportEntry::Indirect { from, name: target }) => {
                            cur_mod = from;
                            cur_name = target;
                        }
                        _ => break,
                    }
                }
                result
            };
            if let Some((_final_mod, _final_name, value)) = terminal {
                out.entry(specifier.clone())
                    .or_default()
                    .insert(name.clone(), value);
            }
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
    host_imports: &crate::compiler::HostImportMetadata,
    modules: &HashMap<String, ModuleRecord>,
) -> Vec<String> {
    let mut unresolved = Vec::new();
    // Imports live on chunk[0] by convention.
    let imports_chunk = match chunks.first() {
        Some(c) => c,
        None => return unresolved,
    };
    for imp in &imports_chunk.imports {
        if imp.module == "*" || imp.module == "env" || imp.module == "wasm:string-constants" {
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

    for import in &host_imports.named {
        let Some(_record) = modules.get(&import.module) else {
            unresolved.push(format!("{}::{}", import.module, import.func));
            continue;
        };
        let mut visited: Vec<(String, String)> = Vec::new();
        if resolve_export(modules, &import.module, &import.func, &mut visited).is_none() {
            unresolved.push(format!("{}::{}", import.module, import.func));
        }
    }

    unresolved.sort();
    unresolved.dedup();
    unresolved
}

/// Recursive resolver — the `ResolveExport(exportName, resolveSet)`
/// abstract op from §16.2.1.6.2. Walks `Indirect` entries until it
/// hits a terminal export or exhausts the chain.
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
        ExportEntry::Function { .. }
        | ExportEntry::Value(_)
        | ExportEntry::ResourceType { .. }
        | ExportEntry::Class { .. } => {
            // Terminal export. Bind to the module that owns this record;
            // runtime installation resolves whether the name materializes
            // as a function marker or immutable value.
            Some((specifier.to_string(), name.to_string()))
        }
        ExportEntry::Indirect {
            from,
            name: src_name,
        } => resolve_export(modules, from, src_name, visited),
    }
}

/// Add shared vybe namespace to all language profiles.
/// This eliminates per-language profile duplication by registering `vybe` as a
/// package-root that gives access to vybe:* modules (gui, types, collections, etc.)
/// Users write: vybe.gui.createForm(), vybe.types.convert(), etc.
fn add_shared_gui_namespace(profile: &mut crate::profile::LanguageProfile) {
    use crate::profile::EsmDefault;

    // Check if `vybe` is already defined as a package-root (shouldn't happen, but be safe)
    let already_has_vybe = profile
        .esm_defaults
        .iter()
        .any(|d| matches!(d, EsmDefault::PackageRoot { prefix, .. } if prefix == "vybe"));

    if !already_has_vybe {
        profile.esm_defaults.push(EsmDefault::PackageRoot {
            prefix: "vybe".to_string(),
            module_root: "vybe".to_string(),
        });
    }
}
