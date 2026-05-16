//! The universal compile unit. A `Bundle` represents everything needed to
//! compile and run — whether loaded from a single source file or a multi-file
//! project. The caller never cares which.

use std::collections::{HashMap, HashSet, VecDeque};
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
    /// Call a static method on a class as the program entry point, e.g. Program.Main().
    Method(String, String),
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
    pub fn prepared_module(&self) -> Result<Module, String> {
        let combined = if self.language.name == "php" {
            expand_php_bundle_sources(&self.sources)?
        } else {
            self.sources.iter()
                .map(|s| s.code.as_str())
                .collect::<Vec<_>>()
                .join("\n")
        };

        // Parse → common AST
        let mut module = (self.language.parse)(&combined)?;

        // Resolve imports relative to first source's directory
        if let Some(first) = self.sources.first() {
            let base_dir = first.path.parent().unwrap_or(Path::new("."));
            resolve_imports(&mut module, &self.language, base_dir);
        }

        // If the project starts with a static method (e.g. C# Program.Main()),
        // inject a call:  ClassName.MethodName()
        if let EntryPoint::Method(ref class_name, ref method_name) = self.entry_point {
            module.body.push(Statement::new(StmtKind::Expr(
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(class_name)),
                        field: method_name.clone(),
                        null_safe: false,
                    })),
                    args: vec![],
                    optional: false,
                })
            )));
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
        let module = self.prepared_module()?;

        // Load profile + compile source code.
        //
        // Flatten the VM's module registry into a per-module map of
        // pre-resolved `(final_module, final_name)` pairs so the
        // compiler Linker can bind `import { X } from "node:http"` in
        // one lookup. Walks the `Indirect` chain from Adapter modules
        // through to a Synthetic `Function` export.
        let mut profile = crate::profile::parse_profile((self.language.profile_source)())?;
        
        // Add shared GUI namespace automatically to all languages
        // (replaces per-language profile duplication)
        add_shared_gui_namespace(&mut profile);
        
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PhpIncludeKind {
    Include,
    IncludeOnce,
    Require,
    RequireOnce,
}

fn expand_php_bundle_sources(sources: &[SourceFile]) -> Result<String, String> {
    let mut included_once = HashSet::new();
    let mut constants: HashMap<String, String> = HashMap::new();
    let mut expanded = Vec::new();
    for source in sources {
        expanded.push(expand_php_source_file(
            &source.path,
            &source.code,
            &mut included_once,
            &mut constants,
        )?);
    }
    Ok(expanded.join("\n"))
}

fn expand_php_source_file(
    source_path: &Path,
    code: &str,
    included_once: &mut HashSet<PathBuf>,
    constants: &mut HashMap<String, String>,
) -> Result<String, String> {
    let mut aliases: HashMap<String, String> = HashMap::new();
    let mut recent_optional_guards: VecDeque<PathBuf> = VecDeque::new();
    let mut brace_depth = 0usize;
    let mut previous_nonempty_line: Option<String> = None;
    let mut out = Vec::new();

    for line in code.lines() {
        let trimmed = line.trim();

        for guarded in extract_php_file_exists_guards(trimmed, source_path, &aliases, constants) {
            recent_optional_guards.push_back(guarded);
            if recent_optional_guards.len() > 8 {
                recent_optional_guards.pop_front();
            }
        }

        if brace_depth == 0
            && let Some((name, value)) = parse_php_alias_assignment(trimmed, source_path, &aliases, constants)
        {
            aliases.insert(name, value);
            out.push(line.to_string());
            brace_depth = update_php_brace_depth(brace_depth, line);
            continue;
        }

        if let Some((name, value)) = parse_php_constant_assignment(trimmed, source_path, &aliases, constants) {
            constants.insert(name, value);
            out.push(line.to_string());
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
                previous_nonempty_line = next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                continue;
            }
            let Some(resolved) = eval_php_path_expression(expr, source_path, &aliases, constants) else {
                if is_dynamic_php_include_expression(expr, constants) {
                    brace_depth = update_php_brace_depth(brace_depth, line);
                    previous_nonempty_line = next_previous_nonempty_line(trimmed, &previous_nonempty_line);
                    continue;
                }
                return Err(format!(
                    "Unsupported PHP include path expression in {}: {}",
                    source_path.display(),
                    trimmed,
                ));
            };
            let include_path = resolve_php_include_path(source_path, &resolved);
            let canonical = std::fs::canonicalize(&include_path).unwrap_or(include_path.clone());

            let should_expand = match kind {
                PhpIncludeKind::IncludeOnce | PhpIncludeKind::RequireOnce => included_once.insert(canonical.clone()),
                PhpIncludeKind::Include | PhpIncludeKind::Require => true,
            };

            if should_expand {
                let include_source = match std::fs::read_to_string(&canonical) {
                    Ok(source) => source,
                    Err(err) => {
                        if err.kind() == std::io::ErrorKind::NotFound
                            && recent_optional_guards.iter().any(|guarded| guarded == &canonical)
                        {
                            continue;
                        }
                        return Err(format!("PHP include error reading {}: {}", canonical.display(), err));
                    }
                };
                let normalized_include = normalize_php_include_for_inlining(&include_source);
                out.push(expand_php_source_file(
                    &canonical,
                    &normalized_include,
                    included_once,
                    constants,
                )?);
            }
            brace_depth = update_php_brace_depth(brace_depth, line);
            continue;
        }

        out.push(normalize_php_alternative_control_syntax(line));
        brace_depth = update_php_brace_depth(brace_depth, line);
        previous_nonempty_line = next_previous_nonempty_line(trimmed, &previous_nonempty_line);
    }

    Ok(out.join("\n"))
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

fn resolve_php_include_path(source_path: &Path, resolved: &str) -> PathBuf {
    let candidate = PathBuf::from(resolved);
    if candidate.is_absolute() {
        candidate
    } else {
        absolutize_path(source_path)
            .parent()
            .unwrap_or(Path::new("."))
            .join(candidate)
    }
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
    if name.is_empty() || !name.chars().all(|ch| ch.is_ascii_alphanumeric() || ch == '_') {
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
    if name.is_empty() || !name.chars().all(|ch| ch.is_ascii_alphanumeric() || ch == '_') {
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
    !parts.is_empty() && parts.into_iter().all(|part| is_dynamic_php_path_atom(part.trim(), constants))
}

fn is_dynamic_php_path_atom(atom: &str, constants: &HashMap<String, String>) -> bool {
    let atom = atom.trim();
    if atom.is_empty() {
        return true;
    }
    if atom.strip_prefix('"').and_then(|s| s.strip_suffix('"')).is_some() {
        return true;
    }
    if atom.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')).is_some() {
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
    if atom.chars().all(|ch| ch.is_ascii_alphanumeric() || ch == '_') {
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

fn next_previous_nonempty_line<'a>(trimmed: &'a str, current: &'a Option<String>) -> Option<String> {
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
        out.push_str(&eval_php_path_atom(part.trim(), source_path, aliases, constants)?);
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
    if atom.chars().all(|ch| ch.is_ascii_alphanumeric() || ch == '_') {
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
            MixedPhpIncludeSegment::Echo { expr, has_close_tag } => {
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
            MixedPhpIncludeSegment::Code { code, has_close_tag } => {
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
            segments.push(MixedPhpIncludeSegment::Echo { expr: code, has_close_tag });
        } else {
            segments.push(MixedPhpIncludeSegment::Code { code, has_close_tag });
        }
        cursor = if has_close_tag { (close + 2).min(source.len()) } else { close };
    }

    if cursor < source.len() {
        segments.push(MixedPhpIncludeSegment::Html(&source[cursor..]));
    }

    segments
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
    let indent_len = line.len() - line.trim_start().len();
    let indent = &line[..indent_len];
    let trimmed = line.trim();
    let Some(inner) = trimmed.strip_prefix("<?php").and_then(|s| s.strip_suffix("?>")) else {
        return line.to_string();
    };
    let inner = inner.trim();

    if let Some(prefix) = inner.strip_suffix(':').map(str::trim_end) {
        if prefix.starts_with("if ")
            || prefix.starts_with("foreach ")
            || prefix.starts_with("for ")
            || prefix.starts_with("while ")
            || prefix.starts_with("switch ")
        {
            return format!("{}<?php {} {{ ?>", indent, prefix);
        }
        if prefix.starts_with("elseif ") {
            return format!("{}<?php }} {} {{ ?>", indent, prefix);
        }
        if prefix == "else" {
            return format!("{}<?php }} else {{ ?>", indent);
        }
    }

    if matches!(inner, "endif;" | "endforeach;" | "endfor;" | "endwhile;" | "endswitch;") {
        return format!("{}<?php }} ?>", indent);
    }

    line.to_string()
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
    use super::*;

    #[test]
    fn php_bundle_expands_require_once_with_dir_alias() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_{}", uuid::Uuid::new_v4()));
        let public_dir = temp_root.join("public");
        std::fs::create_dir_all(&public_dir).expect("create temp dirs");

        let lib_path = temp_root.join("lib.php");
        std::fs::write(&lib_path, "<?php\nfunction shared_helper() { return 1; }\n")
            .expect("write lib");

        let entry_path = public_dir.join("entry.php");
        let entry_src = "<?php\n$basedir = dirname(__DIR__);\nrequire_once $basedir . '/lib.php';\necho shared_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "shared_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_expands_require_once_with_defined_constants() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_define_{}", uuid::Uuid::new_v4()));
        let includes_dir = temp_root.join("includes");
        std::fs::create_dir_all(&includes_dir).expect("create temp dirs");

        let lib_path = includes_dir.join("shared.php");
        std::fs::write(&lib_path, "<?php\nfunction shared_constant_helper() { return 7; }\n")
            .expect("write lib");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\ndefine('WPINC', 'includes');\nrequire_once ABSPATH . WPINC . '/shared.php';\necho shared_constant_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "shared_constant_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_dynamic_variable_include_paths() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_dynamic_include_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction load_dynamic($file) {\n    require_once $file;\n}\nfunction still_present() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "load_dynamic")));
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "still_present")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_parses_define_with_trailing_inline_comment() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_define_comment_{}", uuid::Uuid::new_v4()));
        let includes_dir = temp_root.join("wp-content");
        std::fs::create_dir_all(&includes_dir).expect("create temp dirs");

        let lib_path = includes_dir.join("db-error.php");
        std::fs::write(&lib_path, "<?php\nfunction db_error_helper() { return 9; }\n")
            .expect("write lib");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\ndefine('WP_CONTENT_DIR', ABSPATH . 'wp-content'); // trailing comment\nrequire_once WP_CONTENT_DIR . '/db-error.php';\necho db_error_helper();\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "db_error_helper")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_unknown_constant_include_paths() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_unknown_const_include_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nif (file_exists(WP_CONTENT_DIR . '/db-error.php')) {\n    require_once WP_CONTENT_DIR . '/db-error.php';\n}\nfunction after_optional_include() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_optional_include")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_missing_include_when_guarded_by_file_exists() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_optional_missing_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\nif ( ! file_exists( ABSPATH . '.maintenance' ) ) {\n    return;\n}\nrequire ABSPATH . '.maintenance';\nfunction after_optional_file() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_optional_file")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_function_local_alias_include_paths() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_local_alias_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction load_runtime($engine) {\n    $file = __DIR__ . '/' . $engine . '.php';\n    require_once $file;\n}\nfunction after_runtime_loader() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
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
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_mixed_dynamic_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ndefine('ABSPATH', __DIR__ . '/');\nfunction load_runtime($name) {\n    require_once ABSPATH . 'lib/' . $name . '.php';\n}\nfunction after_mixed_runtime_loader() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "load_runtime")));
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_mixed_runtime_loader")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_does_not_treat_includes_prefixed_function_calls_as_include_statements() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_include_prefix_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction includes_url($path) { return $path; }\n$css = includes_url('style.css');\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_helper_call_include_paths() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_helper_call_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nfunction render_template($base, $file) {\n    require trailingslashit($base) . $file;\n}\nfunction after_helper_include() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_helper_include")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_skips_include_after_missing_positive_defined_guard() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_defined_guard_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\nif ( defined('SUNRISE') ) {\n    include_once __DIR__ . '/sunrise.php';\n}\nfunction after_defined_guard() { return 1; }\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        assert!(module.body.iter().any(|stmt| matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "after_defined_guard")));

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_normalizes_alternative_template_if_syntax() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_alt_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\n$value = true;\n?>\n<?php if ($value): ?>\n<div>ok</div>\n<?php endif; ?>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        bundle.prepared_module().expect("prepared module");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_strips_included_close_tag_before_inline_html() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_close_tag_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let header_path = temp_root.join("header.php");
        std::fs::write(&header_path, "<?php echo 'head'; ?>\n<nav>nav</nav>\n")
            .expect("write header");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ninclude 'header.php';\n?>\n<div>body</div>\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        let echoed_text: String = module.body.iter().filter_map(|stmt| {
            if let StmtKind::Echo(exprs) = &stmt.kind {
                if exprs.len() == 1 {
                    if let ExprKind::Lit(crate::ast::Literal::Str(text)) = &exprs[0].kind {
                        return Some(text.clone());
                    }
                }
            }
            None
        }).collect();
        assert!(!echoed_text.contains("?>"), "bundled inline HTML should not contain a literal close tag: {echoed_text}");

        let _ = std::fs::remove_dir_all(&temp_root);
    }

    #[test]
    fn php_bundle_ignores_boundary_whitespace_from_included_code_file() {
        let temp_root = std::env::temp_dir().join(format!("vybex_php_bundle_boundary_ws_{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&temp_root).expect("create temp dir");

        let helper_path = temp_root.join("helper.php");
        std::fs::write(&helper_path, "\n<?php\nfunction helper_value() { return 1; }\n")
            .expect("write helper");

        let entry_path = temp_root.join("entry.php");
        let entry_src = "<?php\ninclude 'helper.php';\nheader('Location: /next.php');\n";
        std::fs::write(&entry_path, entry_src).expect("write entry");

        let lang = crate::languages::find_by_name("php").expect("php language");
        let bundle = Bundle {
            name: "entry".to_string(),
            language: lang,
            sources: vec![SourceFile { path: entry_path.clone(), code: entry_src.to_string() }],
            wasm_files: vec![],
            entry_point: EntryPoint::Auto,
        };

        let module = bundle.prepared_module().expect("prepared module");
        let echoed_text: String = module.body.iter().filter_map(|stmt| {
            if let StmtKind::Echo(exprs) = &stmt.kind {
                if exprs.len() == 1 {
                    if let ExprKind::Lit(crate::ast::Literal::Str(text)) = &exprs[0].kind {
                        return Some(text.clone());
                    }
                }
            }
            None
        }).collect();
        assert!(echoed_text.is_empty(), "boundary whitespace from included code file should not become output: {echoed_text:?}");

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
        // Non-callable exports — Value (const), ResourceType, Class.
        // `resolve_export` is the host-fn call-target resolver, so
        // class/resource imports don't have a `(module, name)` pair to
        // return here. Constructor calls take a different path that
        // looks up the type_id in the registry and emits a typed
        // construction sequence, not a host call.
        ExportEntry::Value(_)
        | ExportEntry::ResourceType { .. }
        | ExportEntry::Class { .. } => None,
    }
}

/// Add shared vybe namespace to all language profiles.
/// This eliminates per-language profile duplication by registering `vybe` as a 
/// package-root that gives access to vybe:* modules (gui, types, collections, etc.)
/// Users write: vybe.gui.createForm(), vybe.types.convert(), etc.
fn add_shared_gui_namespace(profile: &mut crate::profile::LanguageProfile) {
    use crate::profile::EsmDefault;
    
    // Check if `vybe` is already defined as a package-root (shouldn't happen, but be safe)
    let already_has_vybe = profile.esm_defaults.iter().any(|d| {
        matches!(d, EsmDefault::PackageRoot { prefix, .. } if prefix == "vybe")
    });
    
    if !already_has_vybe {
        profile.esm_defaults.push(EsmDefault::PackageRoot {
            prefix: "vybe".to_string(),
            module_root: "vybe".to_string(),
        });
    }
}

