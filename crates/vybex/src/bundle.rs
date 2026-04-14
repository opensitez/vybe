//! The universal compile unit. A `Bundle` represents everything needed to
//! compile and run — whether loaded from a single source file or a multi-file
//! project. The caller never cares which.

use std::path::{Path, PathBuf};
use crate::ast::*;
use crate::languages::Language;

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

/// Everything needed to compile and run.
pub struct Bundle {
    pub name: String,
    pub language: Language,
    pub sources: Vec<SourceFile>,
    pub entry_point: EntryPoint,
}

impl Bundle {
    /// Parse all sources and compile to bytecode chunks.
    pub fn compile(&self) -> Result<Vec<vybe_bytecode::Chunk>, String> {
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

        // Load profile + compile
        let profile = crate::profile::parse_profile((self.language.profile_source)())?;
        crate::compiler::Compiler::with_profile(profile).compile(&module)
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
        let resolved = base_dir.join(&path_str);
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
