//! Dart walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! (compile_class / compile_expression / etc.) is shared with every
//! other vybex language and works without any Dart-specific knowledge.
//!
//! ## Notes on Dart semantics that the walker normalises
//!
//! - **`this.field` constructor params**: When a constructor has `this.x`,
//!   `this.y` params, we synthesise assignments `this.x = x; this.y = y;`
//!   at the start of the constructor body. The `this.` prefix is stripped
//!   from param names.
//!
//! - **Constructor initializer lists** (`: super(args), field = expr`):
//!   `super(args)` is walked as base_args. `field = expr` assignments are
//!   prepended to the constructor body.
//!
//! - **Factory constructors**: Treated as static methods returning an instance.
//!
//! - **Cascade operator** (`..`): Desugared into a sequence of statements
//!   on the same object using a temp variable pattern.
//!
//! - **Named parameters**: Set `Argument { name: Some(label), value }`.
//!
//! - **`final`/`const` declarations**: Map to `VarDeclKind::Const` (immutable).
//!
//! - **`var`/typed declarations**: Map to `VarDeclKind::Let`.
//!
//! - **Enum declarations**: Each enum value becomes a class constant. Mapped
//!   to `StmtKind::ClassDecl` with `ClassMember::Const` entries.
//!
//! - **For-in**: Always `of: true` — Dart iterates values.
//!
//! - **Switch default**: Emitted as `SwitchCase { conditions: vec![] }` in
//!   source order (not separate `default` field).
//!
//! - **Mixins** (`with Mixin`): Treated as additional parent classes.
//!   Appended to `parents` list after the `extends` parent.

use super::{DartParser, Rule};
use vybe_ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::collections::HashMap;

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs =
        DartParser::parse(Rule::program, source).map_err(|e| format!("Dart parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut mixin_names: std::collections::HashSet<String> = std::collections::HashSet::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI => continue,
            Rule::import_declaration => imports.push(walk_import(pair)?),
            _ => {
                let was_mixin = pair.as_rule() == Rule::mixin_declaration;
                if let Some(stmt) = walk_top_level(pair)? {
                    if was_mixin {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            mixin_names.insert(name.clone());
                        }
                    }
                    body.push(stmt);
                }
            }
        }
    }

    // Mixin merge: copy members from each mixin into classes that
    // declare `with Mixin` (parents). Walker normalisation so the
    // shared class compiler sees a single flat class instead of
    // multi-mixin inheritance.
    apply_mixins(&mut body, &mixin_names);

    Ok(Module {
        name: String::new(),
        language: Lang::Dart,
        body,
        imports,
    })
}

/// Copy methods/fields from each `mixin Foo { ... }` into every
/// `class X with Foo, Bar` (or `class X extends Base with Foo`) and
/// strip the mixin names out of the class's parent list. Mixins
/// themselves stay in the body — they're harmless ClassDecls and
/// some user code may reference them by name.
fn apply_mixins(body: &mut Vec<Statement>, mixin_names: &std::collections::HashSet<String>) {
    if mixin_names.is_empty() {
        return;
    }
    // First pass: collect mixin members keyed by mixin name.
    let mut mixin_members: std::collections::HashMap<String, Vec<ClassMember>> =
        std::collections::HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            if mixin_names.contains(name) {
                mixin_members.insert(name.clone(), members.clone());
            }
        }
    }
    // Second pass: merge into consumers.
    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl {
            name: cname,
            parents,
            members,
            ..
        } = &mut stmt.kind
        {
            if mixin_names.contains(cname) {
                continue;
            }
            let mut new_parents = Vec::new();
            for parent in parents.drain(..) {
                if let Some(mm) = mixin_members.get(&parent) {
                    // Inject mixin members at the END so user-declared
                    // members in the class body win on name conflict.
                    // For Dart's `extends Base with Mixin`, the mixin
                    // override semantics still need work — currently
                    // the class's own methods take priority over the
                    // mixin's, which is the inverse of Dart's "linearization"
                    // rule. Acceptable for simple cases.
                    for m in mm {
                        if !members
                            .iter()
                            .any(|existing| member_name(existing) == member_name(m))
                        {
                            members.push(m.clone());
                        }
                    }
                } else {
                    new_parents.push(parent);
                }
            }
            *parents = new_parents;
        }
    }
}

fn member_name(m: &ClassMember) -> Option<String> {
    match m {
        ClassMember::Field { name, .. } => Some(name.clone()),
        ClassMember::Method(stmt) => match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
            _ => None,
        },
        ClassMember::Const { name, .. } => Some(name.clone()),
        ClassMember::Property { name, .. } => Some(name.clone()),
        _ => None,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Top-level items
// ════════════════════════════════════════════════════════════════════════════

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::mixin_declaration => walk_mixin_decl(pair)?,
        Rule::extension_declaration => walk_extension_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::typedef_declaration => return Ok(None), // type aliases are discarded
        Rule::function_declaration => walk_function_decl(pair)?,
        Rule::variable_declaration_statement => walk_var_decl_stmt(pair)?,
        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().ok_or("empty expr stmt")?)?;
            StmtKind::Expr(expr)
        }
        Rule::annotation => return Ok(None), // annotations discarded at top level
        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

// ════════════════════════════════════════════════════════════════════════════
// Imports
// ════════════════════════════════════════════════════════════════════════════

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut path = String::new();
    let mut alias: Option<String> = None;
    let mut show_names: Vec<ImportName> = Vec::new();
    let mut hide_names: Vec<String> = Vec::new();
    let mut _deferred = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::string_literal => path = unquote_string_literal(&p),
            Rule::as_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_name {
                        alias = Some(c.as_str().to_string());
                    }
                }
            }
            Rule::deferred_clause => {
                _deferred = true;
                for c in p.into_inner() {
                    if c.as_rule() == Rule::as_clause {
                        for a in c.into_inner() {
                            if a.as_rule() == Rule::ident_name {
                                alias = Some(a.as_str().to_string());
                            }
                        }
                    }
                }
            }
            Rule::show_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_list {
                        for name_pair in c.into_inner() {
                            if name_pair.as_rule() == Rule::ident_name {
                                show_names.push(ImportName {
                                    name: name_pair.as_str().to_string(),
                                    alias: None,
                                });
                            }
                        }
                    }
                }
            }
            Rule::hide_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_list {
                        for name_pair in c.into_inner() {
                            if name_pair.as_rule() == Rule::ident_name {
                                hide_names.push(name_pair.as_str().to_string());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    let kind = if !show_names.is_empty() {
        ImportKind::Named {
            path,
            names: show_names,
            level: 0,
        }
    } else {
        ImportKind::Simple { path, alias }
    };

    Ok(Import { kind, span })
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,

        Rule::block_statement => {
            let mut stmts = Vec::new();
            for p in pair.into_inner() {
                if let Some(s) = walk_statement(p)? {
                    stmts.push(s);
                }
            }
            StmtKind::Block(stmts)
        }

        Rule::variable_declaration_statement => walk_var_decl_stmt(pair)?,

        Rule::if_statement => walk_if(pair)?,

        Rule::for_statement => walk_for(pair)?,

        Rule::while_statement => walk_while(pair)?,

        Rule::do_while_statement => walk_do_while(pair)?,

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let expr = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(expr)
        }

        Rule::break_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Break(match label {
                Some(l) => BreakTarget::Label(l),
                None => BreakTarget::Implicit,
            })
        }

        Rule::continue_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Continue(match label {
                Some(l) => ContinueTarget::Label(l),
                None => ContinueTarget::Implicit,
            })
        }

        Rule::throw_statement => {
            let inner = pair.into_inner().next().ok_or("throw: missing expr")?;
            let expr = walk_expression(inner)?;
            StmtKind::Throw {
                expr: Some(expr),
                cause: None,
            }
        }

        Rule::yield_statement => walk_yield_statement(pair)?,

        Rule::rethrow_statement => StmtKind::Throw {
            expr: None,
            cause: None,
        },

        Rule::try_statement => walk_try(pair)?,

        Rule::assert_statement => {
            let mut exprs: Vec<Expression> = Vec::new();
            for p in pair.into_inner() {
                if !is_kw(p.as_rule()) {
                    exprs.push(walk_expression(p)?);
                }
            }
            let test = exprs.remove(0);
            let msg = if exprs.is_empty() {
                None
            } else {
                Some(exprs.remove(0))
            };
            StmtKind::Assert { test, msg }
        }

        Rule::function_declaration => walk_function_decl(pair)?,

        Rule::expression_statement => {
            let inner = pair.into_inner().next().ok_or("empty expr stmt")?;
            let expr = walk_expression(inner)?;
            StmtKind::Expr(expr)
        }

        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,

        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if matches!(
        pair.as_rule(),
        Rule::block_statement | Rule::function_body_block
    ) {
        let mut stmts = Vec::new();
        for p in pair.into_inner() {
            if let Some(s) = walk_statement(p)? {
                stmts.push(s);
            }
        }
        Ok(stmts)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(Vec::new()),
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Variable declarations
// ════════════════════════════════════════════════════════════════════════════

fn walk_var_decl_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // variable_declaration_statement = {
    //     var_modifiers ~ type_or_var ~ var_declarator ~ ("," ~ var_declarator)* ~ ";"
    // }
    let mut var_kind = VarDeclKind::Let;
    let mut declarations = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_modifiers => {
                let txt = p.as_str().trim();
                if txt.contains("final") || txt.contains("const") {
                    var_kind = VarDeclKind::Const;
                }
            }
            Rule::type_or_var => {
                let inner_text = p.as_str().trim();
                if inner_text != "var" {
                    // It's a type annotation, not bare `var`
                    // Check inner children for var_kw
                    let has_var_kw = p.clone().into_inner().any(|c| c.as_rule() == Rule::var_kw);
                    if !has_var_kw {
                        type_hint = Some(inner_text.to_string());
                    }
                }
            }
            Rule::var_declarator => {
                let decl = walk_var_declarator(p, type_hint.clone())?;
                declarations.push(decl);
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind,
    })
}

fn walk_var_declarator(
    pair: Pair<Rule>,
    type_hint: Option<String>,
) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::assignment_expression => init = Some(walk_expression(p)?),
            _ => {
                if init.is_none() {
                    init = Some(walk_expression(p)?);
                }
            }
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint,
        init,
        array_bounds: None,
        with_events: false,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Function declarations
// ════════════════════════════════════════════════════════════════════════════

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type: Option<String> = None;
    let mut is_async = false;
    let mut is_generator = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_annotation => {
                if name.is_empty() {
                    // Return type comes before name
                    return_type = Some(p.as_str().to_string());
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => {} // generic params — discard
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::generator_marker => is_generator = true,
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    is_generator = is_generator || body_has_yield(&body);

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub: false,
    })
}

fn walk_function_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    // function_body = { arrow_body | function_body_block | empty_body }
    let inner = pair.into_inner().next();
    match inner {
        None => Ok(Vec::new()),
        Some(p) => match p.as_rule() {
            Rule::arrow_body => {
                // arrow_body = { "=>" ~ expression ~ ";" }
                let expr_pair = p.into_inner().next().ok_or("arrow body: no expr")?;
                let expr = walk_expression(expr_pair)?;
                Ok(vec![Statement::new(StmtKind::Return(Some(expr)))])
            }
            Rule::function_body_block => walk_statement_into_body(p),
            Rule::empty_body => Ok(Vec::new()),
            _ => Ok(Vec::new()),
        },
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_group => {
                for inner in p.into_inner() {
                    match inner.as_rule() {
                        Rule::param => out.push(walk_param(inner)?),
                        Rule::optional_positional_params => {
                            for op in inner.into_inner() {
                                if op.as_rule() == Rule::param {
                                    let mut param = walk_param(op)?;
                                    param.is_optional = true;
                                    out.push(param);
                                }
                            }
                        }
                        Rule::named_params => {
                            for np in inner.into_inner() {
                                if np.as_rule() == Rule::param {
                                    let mut param = walk_param(np)?;
                                    param.is_optional = true;
                                    out.push(param);
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::param => out.push(walk_param(p)?),
            _ => {}
        }
    }
    Ok(out)
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let mut is_this_param = false;

    fn handle(
        p: Pair<Rule>,
        name: &mut String,
        type_hint: &mut Option<String>,
        default: &mut Option<Expression>,
        is_this: &mut bool,
    ) -> Result<(), String> {
        match p.as_rule() {
            Rule::required_kw | Rule::covariant_kw | Rule::final_kw => {}
            Rule::this_param_prefix => *is_this = true,
            Rule::type_annotation => *type_hint = Some(extract_type_name(&p)),
            Rule::ident_name => *name = p.as_str().to_string(),
            Rule::this_param | Rule::typed_or_untyped_param => {
                // Unwrap the wrapper rule and recurse into its children.
                for inner in p.into_inner() {
                    handle(inner, name, type_hint, default, is_this)?;
                }
            }
            Rule::param_default => {
                let expr_pair = p
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::assignment_expression);
                if let Some(ep) = expr_pair {
                    *default = Some(walk_expression(ep)?);
                }
            }
            _ => {}
        }
        Ok(())
    }
    for p in pair.into_inner() {
        handle(
            p,
            &mut name,
            &mut type_hint,
            &mut default,
            &mut is_this_param,
        )?;
    }

    // this.field params: we keep the bare name. The constructor walker
    // will synthesise `this.name = name;` assignments.
    let _ = is_this_param; // info consumed by constructor_declaration walker

    let is_optional = default.is_some();
    Ok(Param {
        name,
        type_hint,
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    })
}

/// Walk a param and also return whether it was a `this.x` param.
fn walk_param_with_this(pair: Pair<Rule>) -> Result<(Param, bool), String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let mut is_this_param = false;

    fn handle(
        p: Pair<Rule>,
        name: &mut String,
        type_hint: &mut Option<String>,
        default: &mut Option<Expression>,
        is_this: &mut bool,
    ) -> Result<(), String> {
        match p.as_rule() {
            Rule::required_kw | Rule::covariant_kw | Rule::final_kw => {}
            Rule::this_param_prefix => *is_this = true,
            Rule::type_annotation => *type_hint = Some(extract_type_name(&p)),
            Rule::ident_name => *name = p.as_str().to_string(),
            Rule::this_param | Rule::typed_or_untyped_param => {
                for inner in p.into_inner() {
                    handle(inner, name, type_hint, default, is_this)?;
                }
            }
            Rule::param_default => {
                let expr_pair = p
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::assignment_expression);
                if let Some(ep) = expr_pair {
                    *default = Some(walk_expression(ep)?);
                }
            }
            _ => {}
        }
        Ok(())
    }
    for p in pair.into_inner() {
        handle(
            p,
            &mut name,
            &mut type_hint,
            &mut default,
            &mut is_this_param,
        )?;
    }

    let is_optional = default.is_some();
    let param = Param {
        name,
        type_hint,
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    };
    Ok((param, is_this_param))
}

// ════════════════════════════════════════════════════════════════════════════
// Class declarations
// ════════════════════════════════════════════════════════════════════════════

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut modifiers = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::abstract_kw => modifiers.is_abstract = true,
            Rule::class_modifier => {
                if p.into_inner().any(|m| m.as_rule() == Rule::abstract_kw) {
                    modifiers.is_abstract = true;
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => {} // generics — discard
            Rule::extends_clause => {
                if let Some(type_name) = extract_type_name_from_clause(&p) {
                    parents.push(type_name);
                }
            }
            Rule::with_clause => {
                // Mixins become additional parents
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                parents.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::implements_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                interfaces.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::constructor_declaration
                        | Rule::operator_declaration
                        | Rule::getter_declaration
                        | Rule::setter_declaration
                        | Rule::method_declaration
                        | Rule::field_declaration => {
                            if let Some(member) = walk_class_member(m, &name)? {
                                members.push(member);
                            }
                        }
                        Rule::annotation => {} // annotations on members — discard
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    // Static-field rewrite: in static methods, bare `count` (matching
    // a static field name) means `ClassName.count`. Walker rewrites
    // here so the shared compiler doesn't need to track method
    // staticity. Same idea as Fortran's walker normalizing language
    // idioms before they hit the common AST.
    let static_field_names: Vec<String> = members
        .iter()
        .filter_map(|m| {
            if let ClassMember::Field {
                name: fname,
                modifiers,
                ..
            } = m
            {
                if modifiers.is_static {
                    Some(fname.clone())
                } else {
                    None
                }
            } else {
                None
            }
        })
        .collect();
    if !static_field_names.is_empty() {
        for member in members.iter_mut() {
            if let ClassMember::Method(stmt) = member {
                if let StmtKind::FunctionDecl {
                    body, modifiers, ..
                } = &mut stmt.kind
                {
                    if modifiers.is_static {
                        for s in body.iter_mut() {
                            rewrite_static_idents(s, &name, &static_field_names);
                        }
                    }
                }
            }
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers,
        decorators: vec![],
    })
}

/// Rewrite bare `field` to `ClassName.field` for matching static fields.
/// Recursive walk of every Statement / Expression in a static method body.
fn rewrite_static_idents(stmt: &mut Statement, class_name: &str, static_fields: &[String]) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) => {
            rewrite_static_idents_expr(e, class_name, static_fields)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations.iter_mut() {
                if let Some(init) = &mut d.init {
                    rewrite_static_idents_expr(init, class_name, static_fields);
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            for s in then_body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
            for (c, body) in elifs.iter_mut() {
                rewrite_static_idents_expr(c, class_name, static_fields);
                for s in body.iter_mut() {
                    rewrite_static_idents(s, class_name, static_fields);
                }
            }
            if let Some(body) = else_body {
                for s in body.iter_mut() {
                    rewrite_static_idents(s, class_name, static_fields);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init.as_deref_mut() {
                rewrite_static_idents(i, class_name, static_fields);
            }
            if let Some(c) = cond {
                rewrite_static_idents_expr(c, class_name, static_fields);
            }
            if let Some(u) = update {
                rewrite_static_idents_expr(u, class_name, static_fields);
            }
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_static_idents_expr(iter, class_name, static_fields);
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        _ => {}
    }
}

fn rewrite_static_idents_expr(expr: &mut Expression, class_name: &str, static_fields: &[String]) {
    match &mut expr.kind {
        ExprKind::Ident(n) => {
            if static_fields.iter().any(|f| f == n) {
                let name = n.clone();
                expr.kind = ExprKind::Member {
                    object: Box::new(Expression::ident(class_name)),
                    field: name,
                    null_safe: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_static_idents_expr(left, class_name, static_fields);
            rewrite_static_idents_expr(right, class_name, static_fields);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_static_idents_expr(inner, class_name, static_fields)
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_static_idents_expr(callee, class_name, static_fields);
            for a in args.iter_mut() {
                rewrite_static_idents_expr(&mut a.value, class_name, static_fields);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_static_idents_expr(object, class_name, static_fields)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_static_idents_expr(object, class_name, static_fields);
            rewrite_static_idents_expr(index, class_name, static_fields);
        }
        ExprKind::Assign { target, value } => {
            rewrite_static_idents_expr(target, class_name, static_fields);
            rewrite_static_idents_expr(value, class_name, static_fields);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            rewrite_static_idents_expr(then, class_name, static_fields);
            rewrite_static_idents_expr(else_, class_name, static_fields);
        }
        ExprKind::Array(elems) => {
            for e in elems.iter_mut() {
                rewrite_static_idents_expr(&mut e.value, class_name, static_fields);
            }
        }
        ExprKind::Object(props) => {
            for p in props.iter_mut() {
                if let vybe_ast::ObjectProperty::KeyValue { value, .. } = p {
                    rewrite_static_idents_expr(value, class_name, static_fields);
                }
            }
        }
        _ => {}
    }
}

fn walk_mixin_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => {}
            Rule::on_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                parents.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::implements_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                interfaces.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::constructor_declaration
                        | Rule::operator_declaration
                        | Rule::getter_declaration
                        | Rule::setter_declaration
                        | Rule::method_declaration
                        | Rule::field_declaration => {
                            if let Some(member) = walk_class_member(m, &name)? {
                                members.push(member);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_extension_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Extension on Type { members } — treat as a class with the target type as parent
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_annotation => {
                parents.push(extract_type_name(&p));
            }
            Rule::constructor_declaration
            | Rule::operator_declaration
            | Rule::getter_declaration
            | Rule::setter_declaration
            | Rule::method_declaration
            | Rule::field_declaration => {
                if let Some(member) = walk_class_member(p, &name)? {
                    members.push(member);
                }
            }
            _ => {}
        }
    }

    if name.is_empty() {
        name = "__extension__".to_string();
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers {
            is_static: true,
            ..Default::default()
        },
        decorators: vec![],
    })
}

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members: Vec<EnumMember> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut body_members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::enum_values => {
                for vp in p.into_inner() {
                    match vp.as_rule() {
                        Rule::ident_name => {
                            members.push(EnumMember {
                                name: vp.as_str().to_string(),
                                value: None,
                                constructor_args: Vec::new(),
                            });
                        }
                        Rule::enum_value => {
                            let mut value_name = String::new();
                            let mut constructor_args = Vec::new();
                            for inner in vp.into_inner() {
                                match inner.as_rule() {
                                    Rule::ident_name if value_name.is_empty() => {
                                        value_name = inner.as_str().to_string();
                                    }
                                    Rule::argument_list => {
                                        constructor_args = walk_arguments(inner)?
                                            .into_iter()
                                            .map(|arg| arg.value)
                                            .collect();
                                    }
                                    _ => {}
                                }
                            }
                            members.push(EnumMember {
                                name: value_name,
                                value: None,
                                constructor_args,
                            });
                        }
                        _ => {}
                    }
                }
            }
            Rule::enum_clauses => {
                let raw = p.as_str();
                if let Some(idx) = raw.find("implements") {
                    let tail = &raw[idx + "implements".len()..];
                    interfaces.extend(
                        tail.split(',')
                            .map(str::trim)
                            .filter(|name| !name.is_empty())
                            .map(|name| name.to_string()),
                    );
                }
            }
            Rule::class_member => {
                if let Some(member) = walk_class_member(p, &name)? {
                    body_members.push(member);
                }
            }
            Rule::type_params => {}
            _ => {}
        }
    }

    Ok(StmtKind::EnumDecl {
        name,
        interfaces: Vec::new(),
        members,
        visibility: Visibility::Public,
        is_flags: false,
        backing_type: None,
        body_members,
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Class members
// ════════════════════════════════════════════════════════════════════════════

fn walk_class_member(pair: Pair<Rule>, class_name: &str) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::constructor_declaration => Ok(Some(walk_constructor(pair, class_name)?)),
        Rule::method_declaration => Ok(Some(walk_method(pair)?)),
        Rule::field_declaration => walk_field(pair),
        Rule::getter_declaration => Ok(Some(walk_getter(pair)?)),
        Rule::setter_declaration => Ok(Some(walk_setter(pair)?)),
        Rule::operator_declaration => Ok(Some(walk_operator(pair)?)),
        Rule::annotation => Ok(None),
        _ => Ok(None),
    }
}

fn walk_member_modifiers(pair: &Pair<Rule>) -> Modifiers {
    let txt = pair.as_str();
    let mut m = Modifiers::default();
    if txt.contains("static") {
        m.is_static = true;
    }
    if txt.contains("abstract") {
        m.is_abstract = true;
    }
    if txt.contains("override") {
        m.is_override = true;
    }
    if txt.contains("final") {
        m.is_readonly = true;
    }
    m
}

fn walk_constructor(pair: Pair<Rule>, class_name: &str) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args: Option<Vec<Expression>> = None;
    let mut this_params: Vec<String> = Vec::new();
    let mut field_inits: Vec<Statement> = Vec::new();
    let mut is_factory = false;
    let mut _named_ctor: Option<String> = None;
    let mut found_name = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => {
                let txt = p.as_str();
                if txt.contains("factory") {
                    is_factory = true;
                }
            }
            Rule::const_kw => {}
            Rule::factory_kw => is_factory = true,
            Rule::ident_name => {
                if !found_name {
                    found_name = true;
                    // First ident is the class name (or named ctor prefix)
                } else {
                    // Second ident is the named constructor suffix
                    _named_ctor = Some(p.as_str().to_string());
                }
            }
            Rule::param_list => {
                for pg in p.into_inner() {
                    match pg.as_rule() {
                        Rule::param_group => {
                            for inner in pg.into_inner() {
                                match inner.as_rule() {
                                    Rule::param => {
                                        let (param, is_this) = walk_param_with_this(inner)?;
                                        if is_this {
                                            this_params.push(param.name.clone());
                                        }
                                        params.push(param);
                                    }
                                    Rule::optional_positional_params | Rule::named_params => {
                                        for op in inner.into_inner() {
                                            if op.as_rule() == Rule::param {
                                                let (mut param, is_this) =
                                                    walk_param_with_this(op)?;
                                                param.is_optional = true;
                                                if is_this {
                                                    this_params.push(param.name.clone());
                                                }
                                                params.push(param);
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        Rule::param => {
                            let (param, is_this) = walk_param_with_this(pg)?;
                            if is_this {
                                this_params.push(param.name.clone());
                            }
                            params.push(param);
                        }
                        _ => {}
                    }
                }
            }
            Rule::initializer_list => {
                for init in p.into_inner() {
                    if init.as_rule() == Rule::initializer {
                        let inner = init.into_inner().next();
                        if let Some(ini) = inner {
                            match ini.as_rule() {
                                Rule::super_call_initializer => {
                                    let mut args = Vec::new();
                                    for sp in ini.into_inner() {
                                        if sp.as_rule() == Rule::argument_list {
                                            args = walk_arguments(sp)?
                                                .into_iter()
                                                .map(|a| a.value)
                                                .collect();
                                        }
                                    }
                                    base_args = Some(args);
                                }
                                Rule::this_redirect_initializer => {
                                    // `Point.origin() : this(0, 0)` — named
                                    // constructor redirecting to the unnamed
                                    // (or another named) constructor. Walker
                                    // lowers to: `var _self = ClassName(args);
                                    // return _self;` so the named ctor becomes
                                    // a factory-style static method.
                                    let mut redirect_target = None;
                                    let mut redirect_args = Vec::new();
                                    for sp in ini.into_inner() {
                                        match sp.as_rule() {
                                            Rule::ident_name => {
                                                redirect_target = Some(sp.as_str().to_string())
                                            }
                                            Rule::argument_list => {
                                                redirect_args = walk_arguments(sp)?
                                            }
                                            _ => {}
                                        }
                                    }
                                    let new_class = match redirect_target {
                                        Some(name) => Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::ident(class_name)),
                                            field: name,
                                            null_safe: false,
                                        }),
                                        None => Expression::ident(class_name),
                                    };
                                    field_inits.push(Statement::new(StmtKind::Return(Some(
                                        Expression::new(ExprKind::New {
                                            class: Box::new(new_class),
                                            args: redirect_args,
                                        }),
                                    ))));
                                    is_factory = true;
                                }
                                Rule::field_initializer => {
                                    let mut field_name = String::new();
                                    let mut value_expr = None;
                                    for fp in ini.into_inner() {
                                        match fp.as_rule() {
                                            Rule::ident_name => {
                                                field_name = fp.as_str().to_string()
                                            }
                                            Rule::assignment_expression => {
                                                value_expr = Some(walk_expression(fp)?);
                                            }
                                            _ => {}
                                        }
                                    }
                                    if let Some(val) = value_expr {
                                        // Synthesize: this.field = expr;
                                        field_inits.push(Statement::new(StmtKind::Expr(
                                            Expression::new(ExprKind::Assign {
                                                target: Box::new(Expression::new(
                                                    ExprKind::Member {
                                                        object: Box::new(Expression::new(
                                                            ExprKind::This,
                                                        )),
                                                        field: field_name,
                                                        null_safe: false,
                                                    },
                                                )),
                                                value: Box::new(val),
                                            }),
                                        )));
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            Rule::function_body_block => {
                body = walk_statement_into_body(p)?;
            }
            _ => {}
        }
    }

    // Synthesize this.field = field assignments for this.* params
    let mut this_assigns: Vec<Statement> = this_params
        .iter()
        .map(|name| {
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: name.clone(),
                    null_safe: false,
                })),
                value: Box::new(Expression::ident(name)),
            })))
        })
        .collect();

    // Prepend: this.field assignments, then initializer list assignments, then body
    let mut full_body = Vec::new();
    full_body.append(&mut this_assigns);
    full_body.append(&mut field_inits);
    full_body.append(&mut body);

    if is_factory {
        // Factory constructor → static method
        Ok(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: _named_ctor.unwrap_or_else(|| "create".to_string()),
                params,
                return_type: Some(class_name.to_string()),
                body: full_body,
                modifiers: Modifiers {
                    is_static: true,
                    ..Default::default()
                },
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))))
    } else {
        Ok(ClassMember::Constructor {
            params,
            body: full_body,
            base_args,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        })
    }
}

fn walk_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();
    let mut is_async = false;
    let mut is_generator = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {
                if return_type.is_none() {
                    return_type = Some(extract_type_name(&p));
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => {}
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::generator_marker => is_generator = true,
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    is_generator = is_generator || body_has_yield(&body);

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async,
            is_generator,
            is_sub: false,
        },
    ))))
}

fn walk_field(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    // field_declaration = {
    //     member_modifiers ~ type_annotation? ~ ident_name ~ ("=" ~ assignment_expression)?
    //     ~ ("," ~ ident_name ~ ("=" ~ assignment_expression)?)* ~ ";"
    // }
    let mut modifiers = Modifiers::default();
    let mut type_hint: Option<String> = None;
    let mut fields: Vec<(String, Option<Expression>)> = Vec::new();
    let mut current_name = String::new();
    let mut is_const = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => {
                modifiers = walk_member_modifiers(&p);
                let txt = p.as_str();
                if txt.contains("const") {
                    is_const = true;
                }
            }
            Rule::type_annotation => {
                if type_hint.is_none() {
                    type_hint = Some(extract_type_name(&p));
                }
            }
            Rule::ident_name => {
                if !current_name.is_empty() {
                    fields.push((current_name.clone(), None));
                }
                current_name = p.as_str().to_string();
            }
            Rule::assignment_expression => {
                let init = walk_expression(p)?;
                fields.push((current_name.clone(), Some(init)));
                current_name = String::new();
            }
            _ => {}
        }
    }
    if !current_name.is_empty() {
        fields.push((current_name, None));
    }

    // Return first field (most common case: single field)
    // For multiple fields in one declaration, we return the first and
    // the caller should handle multi-field, but our grammar walks them individually.
    if let Some((name, init)) = fields.into_iter().next() {
        if is_const {
            Ok(Some(ClassMember::Const {
                name,
                type_hint: type_hint.clone(),
                value: init.unwrap_or(Expression::null()),
                visibility: modifiers.visibility,
            }))
        } else {
            Ok(Some(ClassMember::Field {
                name,
                type_hint,
                init,
                modifiers,
                with_events: false,
                array_bounds: None,
            }))
        }
    } else {
        Ok(None)
    }
}

fn walk_getter(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => return_type = Some(extract_type_name(&p)),
            Rule::get_keyword => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::async_kw => {}
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    Ok(ClassMember::Property {
        name,
        type_hint: return_type,
        getter: Some(body),
        setter: None,
        is_auto: false,
        modifiers,
    })
}

fn walk_setter(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {}
            Rule::set_keyword => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => {}
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    let param = if let Some(p) = params.into_iter().next() {
        p
    } else {
        Param {
            name: "value".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }
    };

    Ok(ClassMember::Property {
        name,
        type_hint: None,
        getter: None,
        setter: Some(PropertySetter { param, body }),
        is_auto: false,
        modifiers,
    })
}

fn walk_operator(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut op_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();
    let mut return_type: Option<String> = None;
    let mut is_async = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {
                if return_type.is_none() {
                    return_type = Some(extract_type_name(&p));
                }
            }
            Rule::operator_symbol => {
                op_name = match p.as_str().trim() {
                    "+" => "__add__".to_string(),
                    "-" => "__sub__".to_string(),
                    "*" => "__mul__".to_string(),
                    "/" => "__div__".to_string(),
                    "~/" => "__idiv__".to_string(),
                    "%" => "__mod__".to_string(),
                    "==" => "__eq__".to_string(),
                    "!=" => "__ne__".to_string(),
                    "<" => "__lt__".to_string(),
                    ">" => "__gt__".to_string(),
                    "<=" => "__le__".to_string(),
                    ">=" => "__ge__".to_string(),
                    "[]" => "__getitem__".to_string(),
                    "[]=" => "__setitem__".to_string(),
                    "~" => "__bitnot__".to_string(),
                    "&" => "__bitand__".to_string(),
                    "|" => "__bitor__".to_string(),
                    "^" => "__bitxor__".to_string(),
                    "<<" => "__shl__".to_string(),
                    ">>" => "__shr__".to_string(),
                    ">>>" => "__ushr__".to_string(),
                    other => format!("__op_{}__", other),
                };
            }
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: op_name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async,
            is_generator: false,
            is_sub: false,
        },
    ))))
}

// ════════════════════════════════════════════════════════════════════════════
// Control flow
// ════════════════════════════════════════════════════════════════════════════

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond_pair = inner.next().ok_or("if: missing cond")?;
    let mut case_bindings = HashMap::new();
    let cond = if cond_pair.as_rule() == Rule::if_case_condition {
        let mut parts = cond_pair.into_inner();
        let subject = walk_expression(parts.next().ok_or("if-case: missing subject")?)?;
        let mut analysis = analyze_dart_pattern(
            parts
                .find(|p| p.as_rule() == Rule::pattern)
                .ok_or("if-case: missing pattern")?,
            &subject,
        )?;
        if let Some(guard_pair) = parts
            .find(|p| p.as_rule() == Rule::when_guard)
            .and_then(|p| p.into_inner().find(|c| c.as_rule() == Rule::conditional_expression))
        {
            let guard = substitute_pattern_bindings(walk_expression(guard_pair)?, &analysis.bindings);
            analysis.cond = and_expr(analysis.cond, guard);
        }
        case_bindings = analysis.bindings;
        analysis.cond
    } else {
        walk_expression(cond_pair)?
    };
    let then_stmt = inner.next().ok_or("if: missing body")?;
    let mut then_body = walk_statement_into_body(then_stmt)?;
    if !case_bindings.is_empty() {
        for stmt in then_body.iter_mut() {
            substitute_pattern_bindings_stmt(stmt, &case_bindings);
        }
    }

    // else clause
    let else_body = match inner.next() {
        Some(else_pair) => {
            if else_pair.as_rule() == Rule::else_clause {
                let else_stmt = else_pair.into_inner().next().ok_or("else: missing body")?;
                Some(walk_statement_into_body(else_stmt)?)
            } else {
                Some(walk_statement_into_body(else_pair)?)
            }
        }
        None => None,
    };

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

/// Property names that are zero-arg getters in Dart but are bound as
/// value-method emitters in the profile (which only fires on Call).
/// Walker rewrites `expr.<name>` to `expr.<name>()` for these so the
/// dispatch path is uniform.
/// Build a Dart `expr is T` test. For primitive Dart types (int,
/// double, num, String, bool, List, Map, Object) we lower to a
/// `typeof`-style check via REF_TYPEOF so the test works on
/// primitives (which don't carry a `__type` field). Class types
/// fall back to the generic ExprKind::IsType which compares
/// `expr.__type == "T"`.
fn build_is_type(expr: Expression, type_name: &str) -> Expression {
    let trimmed = type_name.trim().trim_end_matches('?');
    let typeof_tag: Option<&str> = match trimmed {
        "int" | "double" | "num" => Some("number"),
        "String" => Some("string"),
        "bool" => Some("boolean"),
        _ => None,
    };
    if let Some(tag) = typeof_tag {
        // Synthesise: `typeof expr === "<tag>"`.
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(tag.into())))),
        });
    }
    Expression::new(ExprKind::IsType {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn is_dart_zero_arg_getter(name: &str) -> bool {
    matches!(
        name,
        "isEmpty" | "isNotEmpty" | "isEven" | "isOdd" | "first" | "last" | "length" | "runes"
    )
}

/// Dart record positional field name `$1`/`$2`/… → its 0-based index. Records
/// are array-backed, so `rec.$1` lowers to `rec[0]`.
fn dart_positional_field_index(name: &str) -> Option<i64> {
    let n: i64 = name.strip_prefix('$')?.parse().ok()?;
    (n >= 1).then_some(n - 1)
}

/// Lower a list comprehension `[for (...) elem]` / `[if (...) elem]` to
/// an IIFE that builds the array imperatively. Walker-only normalization;
/// the compiler sees a regular Call(Lambda, []) on the way out.
fn lower_list_comprehension(elements: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let acc = "__compr_acc";
    let mut body: Vec<Statement> = Vec::new();
    // var __compr_acc = [];
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(acc.to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Array(Vec::new()))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    for el in elements {
        body.push(lower_list_element(el, acc)?);
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Ident(acc.to_string()),
    )))));
    Ok(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn lower_list_element(el: Pair<Rule>, acc: &str) -> Result<Statement, String> {
    let inner = el.into_inner().next().ok_or("empty list element")?;
    match inner.as_rule() {
        Rule::collection_for => {
            // collection_for = "for" "(" for_header ")" list_element
            let mut header_pair = None;
            let mut body_pair = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::for_header => header_pair = Some(p),
                    Rule::list_element => body_pair = Some(p),
                    _ => {}
                }
            }
            let header = header_pair.ok_or("collection_for: missing header")?;
            let body_el = body_pair.ok_or("collection_for: missing body")?;
            let body_stmt = lower_list_element(body_el, acc)?;
            build_for_with_body(header, vec![body_stmt])
        }
        Rule::collection_if => {
            // collection_if = "if" "(" expression ")" list_element ("else" list_element)?
            let mut cond = None;
            let mut then_el = None;
            let mut else_el = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::expression if cond.is_none() => cond = Some(walk_expression(p)?),
                    Rule::list_element => {
                        if then_el.is_none() {
                            then_el = Some(p);
                        } else {
                            else_el = Some(p);
                        }
                    }
                    _ => {}
                }
            }
            let cond = cond.ok_or("collection_if: missing cond")?;
            let then_stmt = lower_list_element(then_el.ok_or("collection_if: missing then")?, acc)?;
            let else_stmt = match else_el {
                Some(el) => Some(vec![lower_list_element(el, acc)?]),
                None => None,
            };
            Ok(Statement::new(StmtKind::If {
                cond,
                then_body: vec![then_stmt],
                elifs: Vec::new(),
                else_body: else_stmt,
            }))
        }
        _ => {
            // Plain expression (or `... ~ expr` spread). Build `acc.add(expr)`.
            // Note: spread is not handled here yet — falls through as a single
            // value push (acceptable for compile_ok; runtime correctness for
            // spread inside comprehensions is a follow-up).
            let value = walk_expression(inner)?;
            let push_call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Ident(acc.to_string()))),
                    field: "add".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(value)],
                optional: false,
            });
            Ok(Statement::new(StmtKind::Expr(push_call)))
        }
    }
}

fn build_for_with_body(header_pair: Pair<Rule>, body: Vec<Statement>) -> Result<Statement, String> {
    let header_inner = header_pair.into_inner().next().ok_or("for: empty header")?;
    match header_inner.as_rule() {
        Rule::for_in_header => {
            let mut var_name = String::new();
            let mut iter_expr = None;
            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::final_kw | Rule::var_kw | Rule::type_annotation => {}
                    Rule::ident_name => var_name = p.as_str().to_string(),
                    _ => iter_expr = Some(walk_expression(p)?),
                }
            }
            Ok(Statement::new(StmtKind::ForIn {
                var: var_name,
                key: None,
                iter: iter_expr.ok_or("for-in: missing iterable")?,
                body,
                of: true,
                else_body: None,
                is_async: false,
            }))
        }
        Rule::for_c_header => {
            let mut init: Option<Box<Statement>> = None;
            let mut cond: Option<Expression> = None;
            let mut update: Option<Expression> = None;
            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("for init: empty")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let stmt_kind = walk_var_decl_no_semi(inner)?;
                                init = Some(Box::new(Statement::new(stmt_kind)));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        if cond.is_none() {
                            cond = Some(walk_expression(p)?);
                        }
                    }
                    Rule::for_c_update => {
                        let exprs: Result<Vec<Expression>, String> =
                            p.into_inner().map(walk_expression).collect();
                        let exprs = exprs?;
                        update = Some(if exprs.len() == 1 {
                            exprs.into_iter().next().unwrap()
                        } else {
                            Expression::new(ExprKind::Sequence(exprs))
                        });
                    }
                    _ => {}
                }
            }
            Ok(Statement::new(StmtKind::For {
                init,
                cond,
                update,
                body,
            }))
        }
        _ => Err(format!(
            "collection_for: unexpected header rule {:?}",
            header_inner.as_rule()
        )),
    }
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // for_statement = { "for" ~ "(" ~ for_header ~ ")" ~ statement }
    let mut header_pair = None;
    let mut body_pair = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_header => header_pair = Some(p),
            _ => body_pair = Some(p),
        }
    }

    let header = header_pair.ok_or("for: missing header")?;
    let body = walk_statement_into_body(body_pair.ok_or("for: missing body")?)?;

    let header_inner = header.into_inner().next().ok_or("for: empty header")?;

    match header_inner.as_rule() {
        Rule::for_in_header => {
            // for_in_header = { (final_kw | var_kw)? ~ type_annotation? ~ ident_name ~ "in" ~ expression }
            let mut var_name = String::new();
            let mut iter_expr = None;
            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::final_kw | Rule::var_kw | Rule::type_annotation => {}
                    Rule::ident_name => var_name = p.as_str().to_string(),
                    _ => iter_expr = Some(walk_expression(p)?),
                }
            }
            Ok(StmtKind::ForIn {
                var: var_name,
                key: None,
                iter: iter_expr.ok_or("for-in: missing iterable")?,
                body,
                of: true, // Dart for-in iterates values
                else_body: None,
                is_async: false,
            })
        }
        Rule::for_c_header => {
            // for_c_header = { for_c_init? ~ ";" ~ expression? ~ ";" ~ for_c_update? }
            let mut init: Option<Box<Statement>> = None;
            let mut cond: Option<Expression> = None;
            let mut update: Option<Expression> = None;

            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("for init: empty")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let stmt_kind = walk_var_decl_no_semi(inner)?;
                                init = Some(Box::new(Statement::new(stmt_kind)));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        if cond.is_none() {
                            cond = Some(walk_expression(p)?);
                        }
                    }
                    Rule::for_c_update => {
                        let exprs: Result<Vec<Expression>, String> =
                            p.into_inner().map(walk_expression).collect();
                        let exprs = exprs?;
                        update = Some(if exprs.len() == 1 {
                            exprs.into_iter().next().unwrap()
                        } else {
                            Expression::new(ExprKind::Sequence(exprs))
                        });
                    }
                    _ => {}
                }
            }

            Ok(StmtKind::For {
                init,
                cond,
                update,
                body,
            })
        }
        other => Err(format!("Unexpected for header: {:?}", other)),
    }
}

fn walk_var_decl_no_semi(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var_kind = VarDeclKind::Let;
    let mut declarations = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_modifiers => {
                let txt = p.as_str().trim();
                if txt.contains("final") || txt.contains("const") {
                    var_kind = VarDeclKind::Const;
                }
            }
            Rule::type_or_var => {
                let inner_text = p.as_str().trim();
                if inner_text != "var" {
                    let has_var_kw = p.clone().into_inner().any(|c| c.as_rule() == Rule::var_kw);
                    if !has_var_kw {
                        type_hint = Some(inner_text.to_string());
                    }
                }
            }
            Rule::var_declarator => {
                let decl = walk_var_declarator(p, type_hint.clone())?;
                declarations.push(decl);
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind,
    })
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("while: missing cond")?)?;
    let body = walk_statement_into_body(inner.next().ok_or("while: missing body")?)?;
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body_pair = inner.next().ok_or("do-while: missing body")?;
    let body = walk_statement_into_body(body_pair)?;
    let cond = walk_expression(inner.next().ok_or("do-while: missing cond")?)?;
    Ok(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(inner.next().ok_or("switch: missing expr")?)?;
    let mut cases = Vec::new();

    for p in inner {
        if p.as_rule() != Rule::switch_case {
            continue;
        }

        let src = p.as_str().trim_start();
        let is_default = src.starts_with("default");

        let mut children: Vec<Pair<Rule>> = p.into_inner().collect();

        if is_default {
            // Default case: all children are body statements
            let stmts = children
                .into_iter()
                .filter_map(|c| walk_statement(c).transpose())
                .collect::<Result<Vec<_>, _>>()?;
            cases.push(SwitchCase {
                conditions: vec![], // empty = default
                body: stmts,
            });
        } else {
            // case expr: stmts...
            // First child that is an expression is the case value
            let case_val = if !children.is_empty() {
                let first = children.remove(0);
                walk_expression(first)?
            } else {
                Expression::null()
            };
            let stmts = children
                .into_iter()
                .filter_map(|c| walk_statement(c).transpose())
                .collect::<Result<Vec<_>, _>>()?;
            cases.push(SwitchCase {
                conditions: vec![CaseCondition::Value(case_val)],
                body: stmts,
            });
        }
    }

    Ok(StmtKind::Switch {
        expr,
        cases,
        default: None,
    })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => {
                if body.is_empty() {
                    body = walk_statement_into_body(p)?;
                }
            }
            Rule::catch_clause => {
                let inner = p.into_inner().next().ok_or("catch: empty")?;
                match inner.as_rule() {
                    Rule::on_catch_clause => {
                        // on Type catch (e, s) { }
                        let mut types = Vec::new();
                        let mut var_name: Option<String> = None;
                        let mut stack_var: Option<String> = None;
                        let mut catch_body = Vec::new();
                        let mut found_first_ident = false;

                        for cp in inner.into_inner() {
                            match cp.as_rule() {
                                Rule::ident_name => {
                                    if types.is_empty() && !found_first_ident {
                                        types.push(cp.as_str().to_string());
                                        found_first_ident = true;
                                    } else if var_name.is_none() {
                                        var_name = Some(cp.as_str().to_string());
                                    } else {
                                        stack_var = Some(cp.as_str().to_string());
                                    }
                                }
                                Rule::block_statement => {
                                    catch_body = walk_statement_into_body(cp)?;
                                }
                                _ => {}
                            }
                        }
                        catches.push(CatchClause {
                            types,
                            var_name,
                            stack_var,
                            body: catch_body,
                            when_clause: None,
                        });
                    }
                    Rule::plain_catch_clause => {
                        // catch (e, s) { }
                        let mut var_name: Option<String> = None;
                        let mut stack_var: Option<String> = None;
                        let mut catch_body = Vec::new();

                        for cp in inner.into_inner() {
                            match cp.as_rule() {
                                Rule::ident_name => {
                                    if var_name.is_none() {
                                        var_name = Some(cp.as_str().to_string());
                                    } else {
                                        stack_var = Some(cp.as_str().to_string());
                                    }
                                }
                                Rule::block_statement => {
                                    catch_body = walk_statement_into_body(cp)?;
                                }
                                _ => {}
                            }
                        }
                        catches.push(CatchClause {
                            types: Vec::new(),
                            var_name,
                            stack_var,
                            body: catch_body,
                            when_clause: None,
                        });
                    }
                    _ => {}
                }
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_statement_into_body(fp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

fn walk_yield_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let is_yield_from = pair.as_str().trim_start().starts_with("yield*");
    let mut value = None;
    for part in pair.into_inner() {
        if !is_kw(part.as_rule()) {
            value = Some(walk_expression(part)?);
        }
    }

    let expr = if is_yield_from {
        ExprKind::YieldFrom(Box::new(value.unwrap_or_else(Expression::null)))
    } else {
        ExprKind::Yield(value.map(Box::new))
    };
    Ok(StmtKind::Expr(Expression::new(expr)))
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // ── Literals ────────────────────────────────────────────────────
        Rule::numeric_literal => {
            let s = pair.as_str().replace('_', "");
            if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(
                    s.parse().map_err(|e| format!("{}", e))?,
                )))
            } else if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?,
                )))
            } else if s.starts_with("0b") || s.starts_with("0B") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?,
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }

        // ── String literals ─────────────────────────────────────────────
        Rule::string_literal => {
            let inner = pair.into_inner().next().ok_or("empty string")?;
            walk_string_literal(inner)
        }

        Rule::raw_string => {
            let s = pair.as_str();
            // r'...' or r"..."
            let inner = if s.starts_with("r'") {
                &s[2..s.len() - 1]
            } else {
                &s[2..s.len() - 1]
            };
            Ok(ExprKind::Lit(Literal::Str(inner.to_string())))
        }

        // Interpolated strings
        Rule::interpolated_double_string | Rule::interpolated_single_string => {
            walk_interpolated_string(pair)
        }

        Rule::triple_double_string | Rule::triple_single_string => walk_interpolated_string(pair),

        // ── Keywords ────────────────────────────────────────────────────
        Rule::this_kw => Ok(ExprKind::This),
        Rule::super_kw => Ok(ExprKind::Super),
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),

        // ── Identifiers ─────────────────────────────────────────────────
        Rule::ident_name => {
            let name = pair.as_str();
            Ok(ExprKind::Ident(name.to_string()))
        }

        Rule::typed_ident => {
            // `Stream<int>` — keep just the identifier; type args are erased.
            let mut inner = pair.into_inner();
            let name_pair = inner.next().ok_or("typed_ident: missing name")?;
            Ok(ExprKind::Ident(name_pair.as_str().to_string()))
        }

        // ── Comma expression ────────────────────────────────────────────
        Rule::expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let exprs: Vec<Expression> = inner
                    .into_iter()
                    .map(walk_expression)
                    .collect::<Result<Vec<_>, _>>()?;
                if exprs.len() == 1 {
                    Ok(exprs.into_iter().next().unwrap().kind)
                } else {
                    Ok(ExprKind::Sequence(exprs))
                }
            }
        }

        // ── Assignment expression ───────────────────────────────────────
        Rule::assignment_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() >= 2 {
                // Check if first child is lambda_expression
                if inner[0].as_rule() == Rule::lambda_expression {
                    return walk_expr_kind(inner.remove(0));
                }
                // conditional_expression ~ (assignment_op ~ assignment_expression)?
                let left = walk_expression(inner.remove(0))?;
                if inner.is_empty() {
                    return Ok(left.kind);
                }
                let op_str = inner.remove(0).as_str().trim();
                let right = walk_expression(inner.remove(0))?;

                if op_str == "=" {
                    Ok(ExprKind::Assign {
                        target: Box::new(left),
                        value: Box::new(right),
                    })
                } else {
                    let op = match op_str {
                        "+=" => CompoundOp::Add,
                        "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul,
                        "/=" => CompoundOp::Div,
                        "~/=" => CompoundOp::IDiv,
                        "%=" => CompoundOp::Mod,
                        "&=" => CompoundOp::BitAnd,
                        "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor,
                        "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr,
                        ">>>=" => CompoundOp::UShr,
                        "??=" => CompoundOp::NullCoalesce,
                        _ => CompoundOp::Add,
                    };
                    Ok(ExprKind::Assign {
                        target: Box::new(left.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: compound_to_binop(op),
                            left: Box::new(left),
                            right: Box::new(right),
                        })),
                    })
                }
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // ── Lambda / arrow function ─────────────────────────────────────
        Rule::lambda_expression => {
            let mut params = Vec::new();
            let mut body = LambdaBody::Expr(Box::new(Expression::null()));
            let mut is_async = false;

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::lambda_params => {
                        for lp in p.into_inner() {
                            match lp.as_rule() {
                                Rule::lambda_param_list => {
                                    for lparam in lp.into_inner() {
                                        if lparam.as_rule() != Rule::lambda_param {
                                            continue;
                                        }
                                        let mut name = String::new();
                                        let mut type_hint: Option<String> = None;
                                        let mut default = None;
                                        for inner in lparam.into_inner() {
                                            match inner.as_rule() {
                                                Rule::ident_name => {
                                                    name = inner.as_str().to_string()
                                                }
                                                Rule::type_annotation => {
                                                    type_hint = Some(extract_type_name(&inner))
                                                }
                                                Rule::param_default => {
                                                    if let Some(ep) = inner.into_inner().find(|c| {
                                                        c.as_rule() == Rule::assignment_expression
                                                    }) {
                                                        default = Some(walk_expression(ep)?);
                                                    }
                                                }
                                                Rule::typed_lambda_param => {
                                                    for ti in inner.into_inner() {
                                                        match ti.as_rule() {
                                                            Rule::ident_name => name = ti.as_str().to_string(),
                                                            Rule::type_annotation => type_hint = Some(extract_type_name(&ti)),
                                                            Rule::param_default => {
                                                                if let Some(ep) = ti.into_inner()
                                                                    .find(|c| c.as_rule() == Rule::assignment_expression) {
                                                                    default = Some(walk_expression(ep)?);
                                                                }
                                                            }
                                                            _ => {}
                                                        }
                                                    }
                                                }
                                                _ => {}
                                            }
                                        }
                                        let is_optional = default.is_some();
                                        params.push(Param {
                                            name,
                                            type_hint,
                                            default,
                                            pass_by: PassBy::Value,
                                            is_rest: false,
                                            is_kwargs: false,
                                            is_optional,
                                            is_nullable: false,
                                        });
                                    }
                                }
                                Rule::ident_name => {
                                    params = vec![Param {
                                        name: lp.as_str().to_string(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    }];
                                }
                                _ => {}
                            }
                        }
                    }
                    Rule::arrow_op => {}
                    Rule::assignment_expression => {
                        body = LambdaBody::Expr(Box::new(walk_expression(p)?));
                    }
                    Rule::function_body_block => {
                        body = LambdaBody::Block(walk_statement_into_body(p)?);
                    }
                    _ => {}
                }
            }

            Ok(ExprKind::Lambda {
                params,
                body,
                is_async,
                captures: Vec::new(),
            })
        }

        // ── Ternary / conditional ───────────────────────────────────────
        Rule::conditional_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() >= 3 {
                // null_coalesce ~ "?" ~ expression ~ ":" ~ expression
                let cond = walk_expression(inner.remove(0))?;
                let then = walk_expression(inner.remove(0))?;
                let else_ = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Ternary {
                    cond: Box::new(cond),
                    then: Box::new(then),
                    else_: Box::new(else_),
                })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // ── Binary expression (flat Pratt) ──────────────────────────────
        //
        // Grammar collapses 12 precedence layers into one `(operand ~
        // (op ~ operand)*)` rule. The walker climbs the resulting flat
        // sequence into a precedence-correct tree using the standard
        // shunting-yard algorithm.
        Rule::null_coalesce_expression => {
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            walk_pratt(inner)
        }

        // ── Relational unit (unary + is/as/relational suffixes) ─────────
        Rule::relational_unit => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }
            let mut left = walk_expression(inner.remove(0))?;
            for p in inner {
                if p.as_rule() != Rule::relational_suffix {
                    continue;
                }
                let mut children: Vec<Pair<Rule>> = p.into_inner().collect();
                let first = children.remove(0);
                match first.as_rule() {
                    Rule::is_test => {
                        let type_name = extract_type_from_inner(first);
                        left = build_is_type(left, &type_name);
                    }
                    Rule::is_not_test => {
                        let type_name = extract_type_from_inner(first);
                        left = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(build_is_type(left, &type_name)),
                        });
                    }
                    Rule::as_cast => {
                        let type_name = extract_type_from_inner(first);
                        left = Expression::new(ExprKind::Cast {
                            expr: Box::new(left),
                            type_name,
                        });
                    }
                    Rule::relational_op => {
                        let op_str = first.as_str().trim();
                        let right = walk_expression(children.remove(0))?;
                        let op = match op_str {
                            "<" => BinOp::Lt,
                            ">" => BinOp::Gt,
                            "<=" => BinOp::LtEq,
                            ">=" => BinOp::GtEq,
                            _ => BinOp::Lt,
                        };
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        });
                    }
                    _ => {}
                }
            }
            Ok(left.kind)
        }

        // ── Unary ───────────────────────────────────────────────────────
        Rule::unary_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }
            // unary_op ~ unary_expression
            let first = inner.remove(0);
            if first.as_rule() == Rule::unary_op {
                let op_str = first.as_str().trim();
                let operand = walk_expression(inner.remove(0))?;
                if op_str.starts_with("await") {
                    return Ok(ExprKind::Await(Box::new(operand)));
                }
                let op = match op_str {
                    "-" => UnaryOp::Neg,
                    "!" => UnaryOp::Not,
                    "~" => UnaryOp::BitNot,
                    "++" => UnaryOp::PreInc,
                    "--" => UnaryOp::PreDec,
                    _ => UnaryOp::Neg,
                };
                Ok(ExprKind::Unary {
                    op,
                    expr: Box::new(operand),
                })
            } else {
                // postfix_expression fallthrough
                walk_expr_kind(first)
            }
        }

        Rule::unary_op => {
            // Should not be reached directly
            Err(format!("unexpected bare unary_op: {}", pair.as_str()))
        }

        // ── Postfix ─────────────────────────────────────────────────────
        Rule::postfix_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let base = walk_expression(inner.remove(0))?;
            if let Some(postfix) = inner.iter().find(|p| p.as_rule() == Rule::postfix_op) {
                let op = match postfix.as_str() {
                    "++" => UnaryOp::PostInc,
                    "--" => UnaryOp::PostDec,
                    _ => return Ok(base.kind),
                };
                Ok(ExprKind::Unary {
                    op,
                    expr: Box::new(base),
                })
            } else {
                Ok(base.kind)
            }
        }

        // ── Call / member / index chain ─────────────────────────────────
        Rule::call_expression => walk_call_chain(pair),

        Rule::new_expression => {
            let inner = pair.into_inner();
            // new_expression = { "new" ~ ident_name ~ ("." ~ ident_name)? ~ type_args? ~ "(" ~ argument_list? ~ ")" }
            let mut class_parts: Vec<String> = Vec::new();
            let mut args = Vec::new();
            for p in inner {
                match p.as_rule() {
                    Rule::ident_name => class_parts.push(p.as_str().to_string()),
                    Rule::type_args => {}
                    Rule::argument_list => args = walk_arguments(p)?,
                    _ => {}
                }
            }
            let class_name = class_parts.join(".");
            Ok(ExprKind::New {
                class: Box::new(Expression::ident(&class_name)),
                args,
            })
        }

        Rule::const_expression => {
            // const ClassName(args) — treat same as new
            let mut class_parts: Vec<String> = Vec::new();
            let mut args = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::ident_name => class_parts.push(p.as_str().to_string()),
                    Rule::type_args => {}
                    Rule::argument_list => args = walk_arguments(p)?,
                    Rule::const_kw => {}
                    _ => {}
                }
            }
            let class_name = class_parts.join(".");
            Ok(ExprKind::New {
                class: Box::new(Expression::ident(&class_name)),
                args,
            })
        }

        // ── Primary ─────────────────────────────────────────────────────
        Rule::primary => {
            let inner = pair.into_inner().next().ok_or("empty primary")?;
            walk_expr_kind(inner)
        }

        // ── Switch expression (Dart 3) ──────────────────────────────────
        Rule::switch_expression => {
            let mut inner = pair.into_inner();
            let subject = walk_expression(inner.next().ok_or("switch expr: missing subject")?)?;
            let mut arms: Vec<(Expression, Expression)> = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::switch_expr_case {
                    let mut case_inner = p.into_inner();
                    let pattern = case_inner
                        .next()
                        .ok_or("switch expr: missing pattern")?;
                    let mut analysis = analyze_dart_pattern(pattern, &subject)?;
                    let mut body_expr = None;
                    for cp in case_inner {
                        match cp.as_rule() {
                            Rule::when_guard => {
                                if let Some(guard_pair) = cp
                                    .into_inner()
                                    .find(|p| p.as_rule() == Rule::conditional_expression)
                                {
                                    let guard = substitute_pattern_bindings(
                                        walk_expression(guard_pair)?,
                                        &analysis.bindings,
                                    );
                                    analysis.cond = and_expr(analysis.cond, guard);
                                }
                            }
                            Rule::assignment_expression => {
                                body_expr = Some(substitute_pattern_bindings(
                                    walk_expression(cp)?,
                                    &analysis.bindings,
                                ));
                            }
                            _ => {}
                        }
                    }
                    if let Some(body) = body_expr {
                        arms.push((analysis.cond, body));
                    }
                }
            }
            Ok(lower_switch_expr_arms(arms).kind)
        }

        // ── Paren / record expression ───────────────────────────────────
        Rule::record_or_paren => {
            // A lone trailing comma is what makes `(99,)` a one-field record
            // rather than the grouping `(99)`; the comma is dropped by the
            // grammar, so read it off the source.
            let single_field_record = pair.as_str().trim_end().trim_end_matches(')').trim_end().ends_with(',');
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                // () — empty tuple/record
                return Ok(ExprKind::Tuple(Vec::new()));
            }
            // record_or_paren_inner contains record_fields
            let ropi = inner.into_iter().next().unwrap();
            if ropi.as_rule() == Rule::record_or_paren_inner {
                let fields: Vec<Pair<Rule>> = ropi
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::record_field)
                    .collect();

                // Check if any field has a name label (record) or single expression (paren)
                if fields.len() == 1 {
                    let field_children: Vec<Pair<Rule>> =
                        fields.into_iter().next().unwrap().into_inner().collect();
                    if field_children.len() == 1 {
                        let value = walk_expression(field_children.into_iter().next().unwrap())?;
                        // `(x,)` is a one-element record; `(x)` is just grouping.
                        return Ok(if single_field_record {
                            ExprKind::Tuple(vec![value])
                        } else {
                            value.kind
                        });
                    } else if field_children.len() == 2 {
                        // Single named field record `(name: value)` — the shared
                        // canonical named-tuple shape (array-backed + by-name key).
                        let name = field_children[0].as_str().to_string();
                        let value = walk_expression(field_children.into_iter().nth(1).unwrap())?;
                        return Ok(vybe_emitter::tuples::build_named_tuple(vec![(
                            Some(name),
                            value,
                        )]));
                    }
                    // Fallthrough — treat as empty
                    return Ok(ExprKind::Lit(Literal::Null));
                }

                // Multiple fields — could be record or tuple
                let has_named = fields.iter().any(|f| f.clone().into_inner().count() > 1);
                if has_named {
                    // Mixed/named record → the shared canonical named-tuple shape:
                    // array-backed (so `.$1`/`.$2` index positionally) with a
                    // by-name key per labelled field (`.host`, `.port`). One value
                    // across languages (Python namedtuple / C# ValueTuple).
                    let mut record_fields: Vec<(Option<String>, Expression)> = Vec::new();
                    for f in fields {
                        let mut fi = f.into_inner();
                        let first = fi.next().unwrap();
                        if let Some(second) = fi.next() {
                            let key = first.as_str().to_string();
                            record_fields.push((Some(key), walk_expression(second)?));
                        } else {
                            record_fields.push((None, walk_expression(first)?));
                        }
                    }
                    Ok(vybe_emitter::tuples::build_named_tuple(record_fields))
                } else {
                    let exprs: Vec<Expression> = fields
                        .into_iter()
                        .map(|f| walk_expression(f.into_inner().next().unwrap()))
                        .collect::<Result<Vec<_>, _>>()?;
                    if exprs.len() == 1 {
                        Ok(exprs.into_iter().next().unwrap().kind)
                    } else {
                        Ok(ExprKind::Tuple(exprs))
                    }
                }
            } else {
                walk_expr_kind(ropi)
            }
        }

        // ── List literal ────────────────────────────────────────────────
        Rule::list_literal => {
            // If any element is a collection-for or collection-if, lower the
            // whole list to an IIFE that builds it imperatively. Otherwise
            // emit a plain array literal.
            let elements: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::list_element)
                .collect();
            let has_comprehension = elements.iter().any(|p| {
                p.clone()
                    .into_inner()
                    .next()
                    .map(|c| matches!(c.as_rule(), Rule::collection_for | Rule::collection_if))
                    .unwrap_or(false)
            });
            if has_comprehension {
                return Ok(lower_list_comprehension(elements)?);
            }
            let mut out = Vec::new();
            for p in elements {
                let src = p.as_str().trim_start();
                let spread = src.starts_with("...");
                let inner = p
                    .into_inner()
                    .next()
                    .ok_or("empty list element".to_string())?;
                let value = walk_expression(inner)?;
                out.push(ArrayElement {
                    key: None,
                    value,
                    spread,
                    by_ref: false,
                });
            }
            Ok(ExprKind::Array(out))
        }

        // ── Map / set literal ───────────────────────────────────────────
        Rule::map_or_set_literal => {
            let mut props = Vec::new();
            let mut is_set = false;
            let mut is_map = false;

            fn walk_one(
                elem: Pair<Rule>,
                props: &mut Vec<ObjectProperty>,
                is_map: &mut bool,
                is_set: &mut bool,
            ) -> Result<(), String> {
                match elem.as_rule() {
                    Rule::map_or_set_element => {
                        for inner in elem.into_inner() {
                            walk_one(inner, props, is_map, is_set)?;
                        }
                    }
                    Rule::map_entry => {
                        *is_map = true;
                        let mut ei = elem.into_inner();
                        let key = walk_expression(ei.next().ok_or("map entry: no key")?)?;
                        let value = walk_expression(ei.next().ok_or("map entry: no value")?)?;
                        props.push(ObjectProperty::KeyValue { key, value });
                    }
                    Rule::assignment_expression => {
                        *is_set = true;
                        let value = walk_expression(elem)?;
                        props.push(ObjectProperty::KeyValue {
                            key: Expression::null(),
                            value,
                        });
                    }
                    // map_collection_if/for / spread — skip body for now;
                    // grammar tolerates them so source compiles.
                    _ => {}
                }
                Ok(())
            }

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::type_args => {}
                    Rule::map_or_set_body => {
                        for entry in p.into_inner() {
                            walk_one(entry, &mut props, &mut is_map, &mut is_set)?;
                        }
                    }
                    _ => {}
                }
            }

            if is_set && !is_map {
                // Set literal — emit as array for now
                let elements: Vec<ArrayElement> = props
                    .into_iter()
                    .filter_map(|p| match p {
                        ObjectProperty::KeyValue { value, .. } => Some(ArrayElement {
                            key: None,
                            value,
                            spread: false,
                            by_ref: false,
                        }),
                        _ => None,
                    })
                    .collect();
                Ok(ExprKind::Array(elements))
            } else {
                Ok(ExprKind::Object(props))
            }
        }

        // ── Passthrough wrappers ────────────────────────────────────────
        Rule::call_chain => {
            let inner = pair.into_inner().next().ok_or("empty call_chain")?;
            walk_expr_kind(inner)
        }

        other => Err(format!(
            "Dart walker: unexpected expression rule: {:?} = {:?}",
            other,
            pair.as_str()
        )),
    }
}

struct DartPatternAnalysis {
    cond: Expression,
    bindings: HashMap<String, Expression>,
}

fn lower_switch_expr_arms(arms: Vec<(Expression, Expression)>) -> Expression {
    let mut fallback = Expression::null();
    for (cond, body) in arms.into_iter().rev() {
        fallback = Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(body),
            else_: Box::new(fallback),
        });
    }
    fallback
}

fn analyze_dart_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    match pair.as_rule() {
        Rule::pattern => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("pattern: empty")?;
            let mut acc = analyze_dart_pattern(first, subject)?;
            while let Some(op) = inner.next() {
                let rhs_pair = inner.next().ok_or("pattern: missing rhs")?;
                let rhs = analyze_dart_pattern(rhs_pair, subject)?;
                acc.cond = match op.as_str() {
                    "&&" => and_expr(acc.cond, rhs.cond),
                    _ => or_expr(acc.cond, rhs.cond),
                };
                acc.bindings.extend(rhs.bindings);
            }
            Ok(acc)
        }
        Rule::primary_pattern => {
            let inner = pair.into_inner().next().ok_or("primary pattern: empty")?;
            analyze_dart_pattern(inner, subject)
        }
        Rule::wildcard_pattern => Ok(pattern_cond(Expression::bool(true))),
        Rule::variable_pattern => {
            let mut bindings = HashMap::new();
            if let Some(name) = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
            {
                if name != "_" {
                    bindings.insert(name, subject.clone());
                }
            }
            Ok(DartPatternAnalysis {
                cond: Expression::bool(true),
                bindings,
            })
        }
        Rule::null_pattern => Ok(pattern_cond(eq_expr(subject.clone(), Expression::null()))),
        Rule::bool_pattern => {
            let value = pair.as_str().trim() == "true";
            Ok(pattern_cond(eq_expr(subject.clone(), Expression::bool(value))))
        }
        Rule::constant_pattern => analyze_constant_pattern(pair, subject),
        Rule::signed_numeric_pattern => {
            let n = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::numeric_literal)
                .ok_or("signed numeric pattern: missing literal")?;
            let lit = Expression::new(walk_expr_kind(n)?);
            Ok(pattern_cond(eq_expr(
                subject.clone(),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(lit),
                }),
            )))
        }
        Rule::relational_pattern => {
            let op_src = pair.as_str().trim_start();
            let op = if op_src.starts_with("<=") {
                BinOp::LtEq
            } else if op_src.starts_with(">=") {
                BinOp::GtEq
            } else if op_src.starts_with("==") {
                BinOp::Eq
            } else if op_src.starts_with("!=") {
                BinOp::NotEq
            } else if op_src.starts_with('<') {
                BinOp::Lt
            } else {
                BinOp::Gt
            };
            let rhs = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::assignment_expression)
                .map(walk_expression)
                .transpose()?
                .unwrap_or_else(Expression::null);
            Ok(pattern_cond(Expression::new(ExprKind::Binary {
                op,
                left: Box::new(subject.clone()),
                right: Box::new(rhs),
            })))
        }
        Rule::list_pattern => analyze_list_pattern(pair, subject),
        Rule::map_pattern => analyze_map_pattern(pair, subject),
        Rule::record_pattern => analyze_record_pattern(pair, subject),
        Rule::object_pattern => analyze_object_pattern(pair, subject),
        _ => Ok(pattern_cond(eq_expr(
            subject.clone(),
            walk_expression(pair)?,
        ))),
    }
}

fn analyze_constant_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if children.len() == 2
        && children.iter().all(|p| p.as_rule() == Rule::ident_name)
        && children[0]
            .as_str()
            .chars()
            .next()
            .map(|ch| ch.is_ascii_uppercase())
            .unwrap_or(false)
    {
        let mut bindings = HashMap::new();
        let name = children[1].as_str().to_string();
        if name != "_" {
            bindings.insert(name, subject.clone());
        }
        return Ok(DartPatternAnalysis {
            cond: Expression::bool(true),
            bindings,
        });
    }
    let value = children
        .into_iter()
        .next()
        .map(|child| {
            if child.as_rule() == Rule::signed_numeric_pattern {
                let n = child
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::numeric_literal)
                    .ok_or("signed numeric pattern: missing literal")?;
                let lit = Expression::new(walk_expr_kind(n)?);
                Ok(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(lit),
                }))
            } else {
                walk_expression(child)
            }
        })
        .transpose()?
        .unwrap_or_else(Expression::null);
    Ok(pattern_cond(eq_expr(subject.clone(), value)))
}

fn analyze_list_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let elements: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::list_pattern_element)
        .collect();
    let rest_pos = elements.iter().position(|p| {
        p.clone()
            .into_inner()
            .next()
            .map(|c| c.as_rule() == Rule::rest_pattern)
            .unwrap_or(false)
    });
    let fixed_count = elements.len() - usize::from(rest_pos.is_some());
    let mut out = pattern_cond(if rest_pos.is_some() {
        cmp_expr(dart_length(subject.clone()), BinOp::GtEq, Expression::int(fixed_count as i64))
    } else {
        eq_expr(dart_length(subject.clone()), Expression::int(fixed_count as i64))
    });

    let mut index = 0usize;
    for elem in elements {
        let child = elem.into_inner().next().ok_or("list pattern: empty element")?;
        if child.as_rule() == Rule::rest_pattern {
            if let Some(name) = child
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
            {
                if name != "_" {
                    out.bindings
                        .insert(name, dart_method_call(subject.clone(), "sublist", vec![Expression::int(index as i64)]));
                }
            }
            continue;
        }
        let item = Expression::new(ExprKind::Index {
            object: Box::new(subject.clone()),
            index: Box::new(Expression::int(index as i64)),
            null_safe: false,
        });
        let part = analyze_dart_pattern(child, &item)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
        index += 1;
    }
    Ok(out)
}

fn analyze_map_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let mut out = pattern_cond(Expression::bool(true));
    for entry in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::map_pattern_entry)
    {
        let mut inner = entry.into_inner();
        let key = walk_expression(inner.next().ok_or("map pattern: missing key")?)?;
        let value_pat = inner.next().ok_or("map pattern: missing value")?;
        out.cond = and_expr(
            out.cond,
            dart_method_call(subject.clone(), "containsKey", vec![key.clone()]),
        );
        let value = Expression::new(ExprKind::Index {
            object: Box::new(subject.clone()),
            index: Box::new(key),
            null_safe: false,
        });
        let part = analyze_dart_pattern(value_pat, &value)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn analyze_record_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let mut out = pattern_cond(Expression::bool(true));
    let mut index = 0usize;
    for field in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::record_pattern_field)
    {
        let children: Vec<Pair<Rule>> = field.into_inner().collect();
        let (target, pat) = if children.len() == 2 {
            (
                Expression::new(ExprKind::Member {
                    object: Box::new(subject.clone()),
                    field: children[0].as_str().to_string(),
                    null_safe: false,
                }),
                children[1].clone(),
            )
        } else {
            let target = Expression::new(ExprKind::Index {
                object: Box::new(subject.clone()),
                index: Box::new(Expression::int(index as i64)),
                null_safe: false,
            });
            index += 1;
            (target, children[0].clone())
        };
        let part = analyze_dart_pattern(pat, &target)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn analyze_object_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let mut out = pattern_cond(Expression::bool(true));
    for field in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::object_pattern_field)
    {
        let mut inner = field.into_inner();
        let name = inner
            .next()
            .ok_or("object pattern: missing field")?
            .as_str()
            .to_string();
        let pat = inner.next().ok_or("object pattern: missing pattern")?;
        let target = Expression::new(ExprKind::Member {
            object: Box::new(subject.clone()),
            field: name,
            null_safe: false,
        });
        let part = analyze_dart_pattern(pat, &target)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn substitute_pattern_bindings(mut expr: Expression, bindings: &HashMap<String, Expression>) -> Expression {
    substitute_pattern_bindings_in_place(&mut expr, bindings);
    expr
}

fn substitute_pattern_bindings_stmt(stmt: &mut Statement, bindings: &HashMap<String, Expression>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) | StmtKind::Throw { expr: Some(expr), .. } => {
            substitute_pattern_bindings_in_place(expr, bindings);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    substitute_pattern_bindings_in_place(init, bindings);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                substitute_pattern_bindings_in_place(target, bindings);
            }
            substitute_pattern_bindings_in_place(value, bindings);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            substitute_pattern_bindings_in_place(target, bindings);
            substitute_pattern_bindings_in_place(value, bindings);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            for s in then_body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
            for (elif_cond, body) in elifs {
                substitute_pattern_bindings_in_place(elif_cond, bindings);
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = else_body {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init.as_deref_mut() {
                substitute_pattern_bindings_stmt(init, bindings);
            }
            if let Some(cond) = cond {
                substitute_pattern_bindings_in_place(cond, bindings);
            }
            if let Some(update) = update {
                substitute_pattern_bindings_in_place(update, bindings);
            }
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            substitute_pattern_bindings_in_place(iter, bindings);
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::Switch { expr, cases, .. } => {
            substitute_pattern_bindings_in_place(expr, bindings);
            for case in cases {
                for condition in &mut case.conditions {
                    if let CaseCondition::Value(value) = condition {
                        substitute_pattern_bindings_in_place(value, bindings);
                    }
                }
                for s in &mut case.body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
            for catch in catches {
                for s in &mut catch.body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = else_body {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = finally {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        _ => {}
    }
}

fn substitute_pattern_bindings_in_place(expr: &mut Expression, bindings: &HashMap<String, Expression>) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(replacement) = bindings.get(name) {
                *expr = replacement.clone();
            }
        }
        ExprKind::Binary { left, right, .. } => {
            substitute_pattern_bindings_in_place(left, bindings);
            substitute_pattern_bindings_in_place(right, bindings);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Cast { expr: inner, .. } => substitute_pattern_bindings_in_place(inner, bindings),
        ExprKind::Ternary { cond, then, else_ } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            substitute_pattern_bindings_in_place(then, bindings);
            substitute_pattern_bindings_in_place(else_, bindings);
        }
        ExprKind::Member { object, .. } => substitute_pattern_bindings_in_place(object, bindings),
        ExprKind::Index { object, index, .. } => {
            substitute_pattern_bindings_in_place(object, bindings);
            substitute_pattern_bindings_in_place(index, bindings);
        }
        ExprKind::Call { callee, args, .. } | ExprKind::New { class: callee, args } => {
            substitute_pattern_bindings_in_place(callee, bindings);
            for arg in args {
                substitute_pattern_bindings_in_place(&mut arg.value, bindings);
            }
        }
        ExprKind::Assign { target, value } => {
            substitute_pattern_bindings_in_place(target, bindings);
            substitute_pattern_bindings_in_place(value, bindings);
        }
        ExprKind::Array(items) => {
            for item in items {
                substitute_pattern_bindings_in_place(&mut item.value, bindings);
                if let Some(key) = &mut item.key {
                    substitute_pattern_bindings_in_place(key, bindings);
                }
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        substitute_pattern_bindings_in_place(key, bindings);
                        substitute_pattern_bindings_in_place(value, bindings);
                    }
                    ObjectProperty::Spread(value) => substitute_pattern_bindings_in_place(value, bindings),
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                substitute_pattern_bindings_in_place(item, bindings);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                substitute_pattern_bindings_in_place(value, bindings);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(value) | InterpolPart::Formatted(value, _) => {
                        substitute_pattern_bindings_in_place(value, bindings)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Match { subject, arms } => {
            substitute_pattern_bindings_in_place(subject, bindings);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        substitute_pattern_bindings_in_place(condition, bindings);
                    }
                }
                substitute_pattern_bindings_in_place(&mut arm.body, bindings);
            }
        }
        _ => {}
    }
}

fn pattern_cond(cond: Expression) -> DartPatternAnalysis {
    DartPatternAnalysis {
        cond,
        bindings: HashMap::new(),
    }
}

fn eq_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::Eq, right)
}

fn cmp_expr(left: Expression, op: BinOp, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn and_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::And, right)
}

fn or_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::Or, right)
}

fn dart_length(value: Expression) -> Expression {
    dart_method_call(value, "length", Vec::new())
}

fn dart_method_call(object: Expression, name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Binary chain helpers
// ════════════════════════════════════════════════════════════════════════════

/// Walk a binary chain where the operator is implicit (same token repeated).
/// E.g. null_coalesce_expression = { logical_or ~ ("??" ~ logical_or)* }
/// Shunting-yard climber over a flat `(operand ~ (op ~ operand)*)`
/// pair sequence. Reproduces Dart's precedence and right-associativity
/// for `??`. All other operators left-associate.
fn walk_pratt(inner: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    if inner.len() == 1 {
        let mut v = inner;
        return walk_expr_kind(v.remove(0));
    }

    // Operands are at even indices, operators at odd indices.
    let mut output: Vec<Expression> = Vec::new();
    let mut ops: Vec<(BinOp, u8)> = Vec::new();

    let mut iter = inner.into_iter();
    let first = iter.next().ok_or("empty pratt expression")?;
    output.push(walk_expression(first)?);

    let mut buf = iter.collect::<Vec<_>>();
    let mut idx = 0;
    while idx < buf.len() {
        let op_pair = buf[idx].clone();
        let op_str = op_pair.as_str().trim();
        let bin_op = str_to_binop(op_str);
        let prec = pratt_precedence(&bin_op);
        let right_assoc = matches!(bin_op, BinOp::NullCoalesce);

        // Reduce while there's an op on stack with higher (or equal,
        // for left-assoc) precedence.
        while let Some(&(_, top_prec)) = ops.last() {
            let should_pop = if right_assoc {
                top_prec > prec
            } else {
                top_prec >= prec
            };
            if !should_pop {
                break;
            }
            let (top_op, _) = ops.pop().unwrap();
            let right = output.pop().ok_or("pratt: missing right")?;
            let left = output.pop().ok_or("pratt: missing left")?;
            output.push(Expression::new(ExprKind::Binary {
                op: top_op,
                left: Box::new(left),
                right: Box::new(right),
            }));
        }

        ops.push((bin_op, prec));
        idx += 1;
        let operand_pair = buf[idx].clone();
        output.push(walk_expression(operand_pair)?);
        idx += 1;
        let _ = &mut buf; // keep variable live
    }

    while let Some((op, _)) = ops.pop() {
        let right = output.pop().ok_or("pratt: missing right")?;
        let left = output.pop().ok_or("pratt: missing left")?;
        output.push(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        }));
    }

    output
        .pop()
        .map(|e| e.kind)
        .ok_or_else(|| "pratt: empty result".to_string())
}

/// Dart precedence table (higher = tighter binding). Mirrors
/// dart:core operator precedence; tweaks: `??` is right-assoc per
/// spec.
fn pratt_precedence(op: &BinOp) -> u8 {
    match op {
        BinOp::NullCoalesce => 1,
        BinOp::Or => 2,
        BinOp::And => 3,
        BinOp::BitOr => 4,
        BinOp::BitXor => 5,
        BinOp::BitAnd => 6,
        BinOp::Eq | BinOp::NotEq => 7,
        BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq => 8,
        BinOp::Shl | BinOp::Shr | BinOp::UShr => 9,
        BinOp::Add | BinOp::Sub => 10,
        BinOp::Mul | BinOp::Div | BinOp::IDiv | BinOp::Mod => 11,
        _ => 0,
    }
}

fn str_to_binop(op: &str) -> BinOp {
    match op {
        "??" => BinOp::NullCoalesce,
        "||" => BinOp::Or,
        "&&" => BinOp::And,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        ">>>" => BinOp::UShr,
        ">>" => BinOp::Shr,
        "<<" => BinOp::Shl,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "~/" => BinOp::IDiv,
        "%" => BinOp::Mod,
        _ => BinOp::Add,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Call chain walker
// ════════════════════════════════════════════════════════════════════════════

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty call expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain {
            continue;
        }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_inner.is_empty() {
            continue;
        }

        let first_rule = chain_inner[0].as_rule();

        match first_rule {
            Rule::cascade_chain => {
                expr = walk_cascade(expr, chain_inner)?;
            }
            Rule::null_safe_member_access => {
                let nsa = chain_inner.into_iter().next().unwrap();
                // Detect a trailing `(...)` from the raw source — pest
                // doesn't yield a pair for an empty `()`, so we have to
                // look at the substring. Without this `obj?.method()`
                // would be walked as `obj?.method` and the call lost.
                let raw = nsa.as_str();
                let has_call = raw.contains('(');
                let mut name = String::new();
                let mut call_args: Option<Vec<Argument>> = None;
                for p in nsa.into_inner() {
                    match p.as_rule() {
                        Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                        Rule::argument_list => call_args = Some(walk_arguments(p)?),
                        _ => {}
                    }
                }
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name.clone(),
                    null_safe: true,
                });
                if let Some(args) = call_args {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args,
                        optional: false,
                    });
                } else if has_call {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: Vec::new(),
                        optional: false,
                    });
                }
            }
            Rule::member_access => {
                let ma = chain_inner.into_iter().next().unwrap();
                // Detect a trailing `(...)` from the raw source — pest
                // doesn't yield a pair for an empty `()`, so we have to
                // look at the substring. Without this `obj.method()` is
                // walked as `obj.method` (Member only) and the call is
                // silently dropped.
                let raw = ma.as_str();
                let has_call = raw.contains('(');
                let mut name = String::new();
                let mut call_args: Option<Vec<Argument>> = None;
                for p in ma.into_inner() {
                    match p.as_rule() {
                        Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                        Rule::argument_list => call_args = Some(walk_arguments(p)?),
                        _ => {}
                    }
                }
                // Dart `arr.fold(initial, combine)` → `arr.reduce(combine, initial)`
                // — JS-shape, args reversed. Walker normalisation so the
                // shared `__array_reduce` HOF dispatch can handle it.
                if name == "fold" {
                    if let Some(ref mut args) = call_args {
                        if args.len() == 2 {
                            args.swap(0, 1);
                            name = "reduce".to_string();
                        }
                    }
                }
                if let Some(uri) = dart_uri_from_expr(&expr) {
                    if name == "normalizePath" && (has_call || call_args.is_some()) {
                        expr = dart_uri_expr(uri.normalize_path());
                        continue;
                    }
                    if name == "replace" && (has_call || call_args.is_some()) {
                        expr = dart_uri_expr(uri.replace_with(call_args.clone().unwrap_or_default()));
                        continue;
                    }
                    if name == "resolve" {
                        if let Some(args) = &call_args {
                            if let Some(rel) = args.first().and_then(|a| literal_string(&a.value)) {
                                expr = dart_uri_expr(uri.resolve(&rel));
                                continue;
                            }
                        }
                    }
                    if name == "resolveUri" {
                        if let Some(args) = &call_args {
                            if let Some(rel) = args.first().and_then(|a| dart_uri_from_expr(&a.value)) {
                                expr = dart_uri_expr(uri.resolve(&rel.href));
                                continue;
                            }
                        }
                    }
                }
                if let ExprKind::Ident(class_name) = expr.kind.clone() {
                    if matches!(class_name.as_str(), "DateTime" | "Duration" | "Uri" | "List") {
                        expr = Expression::ident(&format!("{}.{}", class_name, name));
                        if class_name == "Duration" && name == "zero" {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args: Vec::new(),
                                optional: false,
                            });
                        }
                        continue;
                    }
                }
                // Dart zero-arg getters that map to value-method emitters
                // need to look like Calls so the value-method dispatch
                // kicks in. Wrap the bare property access in a Call(0)
                // for known property names.
                let force_call = !has_call && call_args.is_none() && is_dart_zero_arg_getter(&name);
                // Dart record positional field `.$1`/`.$2` → indexed read (records
                // are array-backed). Only a bare getter, never a call.
                if let Some(idx) =
                    (call_args.is_none() && !has_call).then(|| dart_positional_field_index(&name)).flatten()
                {
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(Expression::int(idx)),
                        null_safe: false,
                    });
                } else {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: false,
                    });
                    if let Some(args) = call_args {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args,
                            optional: false,
                        });
                    } else if has_call || force_call {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args: Vec::new(),
                            optional: false,
                        });
                    }
                }
            }
            Rule::call_args => {
                let ca = chain_inner.into_iter().next().unwrap();
                let mut args = ca
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::argument_list)
                    .map(walk_arguments)
                    .transpose()?
                    .unwrap_or_default();
                if is_ident_expr(&expr, "RegExp") {
                    args = normalize_regexp_args(args);
                } else if is_ident_expr(&expr, "Duration") {
                    args = normalize_duration_args(args);
                } else if is_ident_expr(&expr, "Uri.parse") {
                    if let Some(text) = args.first().and_then(|arg| literal_string(&arg.value)) {
                        expr = dart_uri_expr(DartUri::parse(&text));
                        continue;
                    }
                } else if is_ident_expr(&expr, "Uri.http") || is_ident_expr(&expr, "Uri.https") {
                    if args.len() >= 2 {
                        if let (Some(authority), Some(path)) =
                            (literal_string(&args[0].value), literal_string(&args[1].value))
                        {
                            let https = is_ident_expr(&expr, "Uri.https");
                            let query = args.get(2).and_then(|arg| query_string_from_expr(&arg.value));
                            let port = args.get(3).and_then(|arg| literal_number_string(&arg.value));
                            expr = dart_uri_expr(DartUri::from_parts(
                                if https { "https" } else { "http" },
                                &authority,
                                &path,
                                query.as_deref(),
                                None,
                                port.as_deref(),
                            ));
                            continue;
                        }
                    }
                } else if is_ident_expr(&expr, "Uri.file") {
                    if let Some(path) = args.first().and_then(|arg| literal_string(&arg.value)) {
                        expr = dart_uri_expr(DartUri::from_file(&path));
                        continue;
                    }
                }
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args,
                    optional: false,
                });
            }
            Rule::index_access => {
                let ia = chain_inner.into_iter().next().unwrap();
                let index_expr = ia
                    .into_inner()
                    .next()
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or(Expression::int(0));
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index_expr),
                    null_safe: false,
                });
            }
            Rule::null_assert => {
                // `!` postfix — null assertion, just pass through
            }
            _ => {
                // Fallback: try to match by source text
                if chain_src.starts_with("?.") {
                    let name = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: true,
                    });
                } else if chain_src.starts_with("(") {
                    let args = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::argument_list)
                        .map(walk_arguments)
                        .transpose()?
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args,
                        optional: false,
                    });
                } else if chain_src.starts_with("[") {
                    let index_expr = chain_inner
                        .into_iter()
                        .find(|p| !matches!(p.as_rule(), Rule::call_chain))
                        .map(walk_expression)
                        .transpose()?
                        .unwrap_or(Expression::int(0));
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(index_expr),
                        null_safe: false,
                    });
                } else if chain_src.starts_with(".") {
                    let name = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: false,
                    });
                }
            }
        }
    }

    Ok(expr.kind)
}

fn is_ident_expr(expr: &Expression, expected: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if name == expected)
}

fn literal_bool(expr: &Expression) -> Option<bool> {
    match &expr.kind {
        ExprKind::Lit(Literal::Bool(value)) => Some(*value),
        _ => None,
    }
}

fn literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        _ => None,
    }
}

fn literal_number_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(value.to_string()),
        ExprKind::Lit(Literal::Float(value)) => Some(format!("{}", value)),
        _ => None,
    }
}

fn normalize_regexp_args(args: Vec<Argument>) -> Vec<Argument> {
    let mut pattern = None;
    let mut case_sensitive = true;
    let mut multi_line = false;
    let mut unicode = false;
    let mut dot_all = false;

    for arg in args {
        match arg.name.as_deref() {
            Some("caseSensitive") => {
                if let Some(value) = literal_bool(&arg.value) {
                    case_sensitive = value;
                }
            }
            Some("multiLine") => {
                if let Some(value) = literal_bool(&arg.value) {
                    multi_line = value;
                }
            }
            Some("unicode") => {
                if let Some(value) = literal_bool(&arg.value) {
                    unicode = value;
                }
            }
            Some("dotAll") => {
                if let Some(value) = literal_bool(&arg.value) {
                    dot_all = value;
                }
            }
            _ if pattern.is_none() => pattern = Some(arg),
            _ => {}
        }
    }

    let mut flags = String::new();
    if !case_sensitive {
        flags.push('i');
    }
    if multi_line {
        flags.push('m');
    }
    if unicode {
        flags.push('u');
    }
    if dot_all {
        flags.push('s');
    }

    let mut out = Vec::new();
    out.push(pattern.unwrap_or(Argument {
        value: Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        name: None,
        by_ref: false,
        spread: false,
    }));
    out.push(Argument {
        value: Expression::new(ExprKind::Lit(Literal::Str(flags))),
        name: None,
        by_ref: false,
        spread: false,
    });
    out
}

fn positional_arg(value: Expression) -> Argument {
    Argument {
        value,
        name: None,
        by_ref: false,
        spread: false,
    }
}

fn mul_expr(value: Expression, factor: f64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(value),
        right: Box::new(Expression::float(factor)),
    })
}

fn add_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn normalize_duration_args(args: Vec<Argument>) -> Vec<Argument> {
    if args.iter().all(|arg| arg.name.is_none()) {
        return args;
    }
    let mut total = Expression::float(0.0);
    for arg in args {
        let factor = match arg.name.as_deref() {
            Some("days") => 86_400_000.0,
            Some("hours") => 3_600_000.0,
            Some("minutes") => 60_000.0,
            Some("seconds") => 1000.0,
            Some("milliseconds") => 1.0,
            Some("microseconds") => 0.001,
            _ => continue,
        };
        total = add_expr(total, mul_expr(arg.value, factor));
    }
    vec![positional_arg(total)]
}

#[derive(Clone, Debug)]
struct DartUri {
    scheme: String,
    host: String,
    port: String,
    path: String,
    query: String,
    fragment: String,
    user_info: String,
    href: String,
}

fn uri_decode(text: &str) -> String {
    text.replace("%20", " ")
}

fn default_port(scheme: &str) -> String {
    match scheme {
        "http" => "80".to_string(),
        "https" => "443".to_string(),
        _ => String::new(),
    }
}

fn split_authority(input: &str) -> (String, String, String) {
    let (user_info, host_port) = input
        .rsplit_once('@')
        .map(|(u, h)| (u.to_string(), h.to_string()))
        .unwrap_or_else(|| (String::new(), input.to_string()));
    if let Some((host, port)) = host_port.rsplit_once(':') {
        if port.chars().all(|c| c.is_ascii_digit()) {
            return (user_info, host.to_string(), port.to_string());
        }
    }
    (user_info, host_port, String::new())
}

fn normalize_path_text(path: &str) -> String {
    let absolute = path.starts_with('/');
    let trailing = path.ends_with('/') && path.len() > 1;
    let mut parts: Vec<&str> = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    let mut out = parts.join("/");
    if absolute {
        out.insert(0, '/');
    }
    if trailing && !out.ends_with('/') {
        out.push('/');
    }
    if out.is_empty() && absolute {
        "/".to_string()
    } else {
        out
    }
}

impl DartUri {
    fn parse(input: &str) -> Self {
        let href = input.to_string();
        let (without_fragment, fragment) = input
            .split_once('#')
            .map(|(a, b)| (a, b))
            .unwrap_or((input, ""));
        let (without_query, query) = without_fragment
            .split_once('?')
            .map(|(a, b)| (a, b))
            .unwrap_or((without_fragment, ""));
        let (scheme, rest) = without_query
            .split_once("://")
            .map(|(s, r)| (s.to_lowercase(), r))
            .unwrap_or_else(|| (String::new(), without_query));
        let raw_path_storage;
        let (authority, raw_path) = if scheme.is_empty() {
            ("", rest)
        } else if let Some((a, p)) = rest.split_once('/') {
            raw_path_storage = format!("/{}", p);
            (a, raw_path_storage.as_str())
        } else {
            (rest, "")
        };
        let (user_info, host, explicit_port) = split_authority(authority);
        let port = if explicit_port.is_empty() {
            default_port(&scheme)
        } else {
            explicit_port
        };
        let path = uri_decode(raw_path);
        Self {
            scheme,
            host,
            port,
            path,
            query: uri_decode(query),
            fragment: uri_decode(fragment),
            user_info: uri_decode(&user_info),
            href,
        }
    }

    fn from_parts(
        scheme: &str,
        authority: &str,
        path: &str,
        query: Option<&str>,
        fragment: Option<&str>,
        port_override: Option<&str>,
    ) -> Self {
        let (user_info, host, explicit_port) = split_authority(authority);
        let port = port_override
            .filter(|p| !p.is_empty() && *p != "null")
            .map(|p| p.to_string())
            .or_else(|| (!explicit_port.is_empty()).then_some(explicit_port))
            .unwrap_or_else(|| default_port(scheme));
        let mut href = format!("{}://{}", scheme, host);
        let default = default_port(scheme);
        if !port.is_empty() && port != default {
            href.push(':');
            href.push_str(&port);
        }
        href.push_str(path);
        if let Some(q) = query.filter(|q| !q.is_empty()) {
            href.push('?');
            href.push_str(q);
        }
        if let Some(f) = fragment.filter(|f| !f.is_empty()) {
            href.push('#');
            href.push_str(f);
        }
        Self {
            scheme: scheme.to_string(),
            host,
            port,
            path: path.to_string(),
            query: query.unwrap_or("").to_string(),
            fragment: fragment.unwrap_or("").to_string(),
            user_info,
            href,
        }
    }

    fn from_file(path: &str) -> Self {
        Self {
            scheme: "file".to_string(),
            host: String::new(),
            port: String::new(),
            path: path.to_string(),
            query: String::new(),
            fragment: String::new(),
            user_info: String::new(),
            href: format!("file://{}", path),
        }
    }

    fn authority(&self) -> String {
        if self.host.is_empty() {
            return String::new();
        }
        let default = default_port(&self.scheme);
        if self.port.is_empty() || self.port == default {
            self.host.clone()
        } else {
            format!("{}:{}", self.host, self.port)
        }
    }

    fn origin(&self) -> String {
        if self.scheme.is_empty() || self.host.is_empty() {
            return String::new();
        }
        format!("{}://{}", self.scheme, self.authority())
    }

    fn normalize_path(mut self) -> Self {
        self.path = normalize_path_text(&self.path);
        self.href = self.recompose();
        self
    }

    fn recompose(&self) -> String {
        if self.scheme.is_empty() {
            let mut out = self.path.clone();
            if !self.query.is_empty() {
                out.push('?');
                out.push_str(&self.query);
            }
            if !self.fragment.is_empty() {
                out.push('#');
                out.push_str(&self.fragment);
            }
            return out;
        }
        let mut out = format!("{}://{}", self.scheme, self.authority());
        out.push_str(&self.path);
        if !self.query.is_empty() {
            out.push('?');
            out.push_str(&self.query);
        }
        if !self.fragment.is_empty() {
            out.push('#');
            out.push_str(&self.fragment);
        }
        out
    }

    fn replace_with(mut self, args: Vec<Argument>) -> Self {
        for arg in args {
            match arg.name.as_deref() {
                Some("scheme") => {
                    if let Some(v) = literal_string(&arg.value) {
                        self.scheme = v;
                    }
                }
                Some("host") => {
                    if let Some(v) = literal_string(&arg.value) {
                        self.host = v;
                    }
                }
                Some("port") => {
                    if let Some(v) = literal_number_string(&arg.value) {
                        self.port = v;
                    }
                }
                Some("path") => {
                    if let Some(v) = literal_string(&arg.value) {
                        self.path = v;
                    }
                }
                Some("query") => {
                    if let Some(v) = literal_string(&arg.value) {
                        self.query = v;
                    }
                }
                Some("fragment") => {
                    if let Some(v) = literal_string(&arg.value) {
                        self.fragment = v;
                    }
                }
                Some("pathSegments") => {
                    if let Some(v) = path_from_segments_expr(&arg.value) {
                        self.path = v;
                    }
                }
                _ => {}
            }
        }
        self.href = self.recompose();
        self
    }

    fn resolve(&self, rel: &str) -> Self {
        if rel.contains("://") {
            return Self::parse(rel);
        }
        let mut out = self.clone();
        if rel.starts_with('/') {
            out.path = rel.to_string();
        } else {
            let base_dir = self
                .path
                .rsplit_once('/')
                .map(|(head, _)| format!("{}/", head))
                .unwrap_or_else(|| "/".to_string());
            out.path = format!("{}{}", base_dir, rel);
        }
        out.path = normalize_path_text(&out.path);
        out.query.clear();
        out.fragment.clear();
        out.href = out.recompose();
        out
    }
}

fn obj_prop(key: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: Expression::string(key),
        value,
    }
}

fn query_params_expr(query: &str) -> Expression {
    let props = query
        .split('&')
        .filter(|part| !part.is_empty())
        .map(|part| {
            let (k, v) = part.split_once('=').unwrap_or((part, ""));
            obj_prop(&uri_decode(k), Expression::string(&uri_decode(v)))
        })
        .collect();
    Expression::new(ExprKind::Object(props))
}

fn path_segments_expr(path: &str) -> Expression {
    Expression::new(ExprKind::Array(
        path.split('/')
            .filter(|part| !part.is_empty())
            .map(|part| ArrayElement {
                value: Expression::string(part),
                key: None,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn dart_uri_expr(uri: DartUri) -> Expression {
    let has_query = !uri.query.is_empty();
    let has_fragment = !uri.fragment.is_empty();
    let has_authority = !uri.host.is_empty();
    Expression::new(ExprKind::Object(vec![
        obj_prop("__dart_uri_marker", Expression::bool(true)),
        obj_prop("scheme", Expression::string(&uri.scheme)),
        obj_prop("host", Expression::string(&uri.host)),
        obj_prop("port", Expression::string(&uri.port)),
        obj_prop("path", Expression::string(&uri.path)),
        obj_prop("query", Expression::string(&uri.query)),
        obj_prop("fragment", Expression::string(&uri.fragment)),
        obj_prop("userInfo", Expression::string(&uri.user_info)),
        obj_prop("authority", Expression::string(&uri.authority())),
        obj_prop("origin", Expression::string(&uri.origin())),
        obj_prop("href", Expression::string(&uri.href)),
        obj_prop("hasScheme", Expression::bool(!uri.scheme.is_empty())),
        obj_prop("hasAuthority", Expression::bool(has_authority)),
        obj_prop("isAbsolute", Expression::bool(!uri.scheme.is_empty())),
        obj_prop("hasQuery", Expression::bool(has_query)),
        obj_prop("hasFragment", Expression::bool(has_fragment)),
        obj_prop("hasEmptyPath", Expression::bool(uri.path.is_empty())),
        obj_prop("pathSegments", path_segments_expr(&uri.path)),
        obj_prop("queryParameters", query_params_expr(&uri.query)),
    ]))
}

fn object_string_field(expr: &Expression, field: &str) -> Option<String> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    props.iter().find_map(|prop| match prop {
        ObjectProperty::KeyValue { key, value } => {
            if literal_string(key).as_deref() == Some(field) {
                literal_string(value)
            } else {
                None
            }
        }
        _ => None,
    })
}

fn dart_uri_from_expr(expr: &Expression) -> Option<DartUri> {
    let ExprKind::Object(_) = &expr.kind else {
        return None;
    };
    if object_string_field(expr, "__dart_uri_marker").is_none()
        && !matches!(
            &expr.kind,
            ExprKind::Object(props) if props.iter().any(|p| matches!(p, ObjectProperty::KeyValue { key, value } if literal_string(key).as_deref() == Some("__dart_uri_marker") && matches!(value.kind, ExprKind::Lit(Literal::Bool(true)))))
        )
    {
        return None;
    }
    Some(DartUri {
        scheme: object_string_field(expr, "scheme").unwrap_or_default(),
        host: object_string_field(expr, "host").unwrap_or_default(),
        port: object_string_field(expr, "port").unwrap_or_default(),
        path: object_string_field(expr, "path").unwrap_or_default(),
        query: object_string_field(expr, "query").unwrap_or_default(),
        fragment: object_string_field(expr, "fragment").unwrap_or_default(),
        user_info: object_string_field(expr, "userInfo").unwrap_or_default(),
        href: object_string_field(expr, "href").unwrap_or_default(),
    })
}

fn query_string_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Null) => Some(String::new()),
        ExprKind::Object(props) => Some(
            props
                .iter()
                .filter_map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        Some(format!("{}={}", literal_string(key)?, literal_string(value)?))
                    }
                    _ => None,
                })
                .collect::<Vec<_>>()
                .join("&"),
        ),
        _ => None,
    }
}

fn path_from_segments_expr(expr: &Expression) -> Option<String> {
    let ExprKind::Array(items) = &expr.kind else {
        return None;
    };
    let parts = items
        .iter()
        .map(|item| literal_string(&item.value))
        .collect::<Option<Vec<_>>>()?;
    Some(format!("/{}", parts.join("/")))
}

// ════════════════════════════════════════════════════════════════════════════
// Cascade desugaring
// ════════════════════════════════════════════════════════════════════════════

/// Desugar `obj..method()..field = val` into a sequence on the same object.
/// We create a block expression pattern by wrapping the cascade into
/// assignments on the receiver.
fn walk_cascade(receiver: Expression, chain_inner: Vec<Pair<Rule>>) -> Result<Expression, String> {
    let cascade_chain = chain_inner.into_iter().next().ok_or("cascade: empty")?;
    let mut sections = Vec::new();

    for p in cascade_chain.into_inner() {
        match p.as_rule() {
            Rule::cascade_op => {} // ".." or "?.."
            Rule::cascade_section => sections.push(p),
            Rule::cascade_continuation => {
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::cascade_section {
                        sections.push(cp);
                    }
                }
            }
            _ => {}
        }
    }

    let mut ops = Vec::new();

    for section in sections {
        let mut sec_inner = section.into_inner();
        let first = sec_inner.next().ok_or("cascade section: empty")?;

        match first.as_rule() {
            Rule::ident_name => {
                let name = first.as_str().to_string();
                // Check what follows: call, assignment, or bare member
                if let Some(next_p) = sec_inner.next() {
                    if next_p.as_rule() == Rule::argument_list {
                        // method call: receiver.name(args)
                        let args = walk_arguments(next_p)?;
                        let callee = Expression::new(ExprKind::Member {
                            object: Box::new(receiver.clone()),
                            field: name,
                            null_safe: false,
                        });
                        ops.push(Expression::new(ExprKind::Call {
                            callee: Box::new(callee),
                            args,
                            optional: false,
                        }));
                    } else {
                        // assignment: receiver.name = expr
                        let value = walk_expression(next_p)?;
                        ops.push(Expression::new(ExprKind::Assign {
                            target: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(receiver.clone()),
                                field: name,
                                null_safe: false,
                            })),
                            value: Box::new(value),
                        }));
                    }
                } else {
                    // Bare member access
                    ops.push(Expression::new(ExprKind::Member {
                        object: Box::new(receiver.clone()),
                        field: name,
                        null_safe: false,
                    }));
                }
            }
            _ => {
                // Index cascade: [expr] or [expr] = expr
                // Just pass through for now
            }
        }
    }

    ops.push(receiver);
    Ok(Expression::new(ExprKind::Sequence(ops)))
}

// ════════════════════════════════════════════════════════════════════════════
// Arguments
// ════════════════════════════════════════════════════════════════════════════

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() != Rule::argument {
            continue;
        }
        let inner = p.into_inner().next().ok_or("empty argument")?;
        match inner.as_rule() {
            Rule::named_argument => {
                let mut name = String::new();
                let mut value = Expression::null();
                for np in inner.into_inner() {
                    match np.as_rule() {
                        Rule::ident_name => name = np.as_str().to_string(),
                        Rule::assignment_expression => value = walk_expression(np)?,
                        _ => {}
                    }
                }
                args.push(Argument {
                    value,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                });
            }
            Rule::spread_argument => {
                let expr_pair = inner.into_inner().next().ok_or("spread: no expr")?;
                let value = walk_expression(expr_pair)?;
                args.push(Argument {
                    value,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            }
            Rule::assignment_expression => {
                let value = walk_expression(inner)?;
                args.push(Argument::positional(value));
            }
            _ => {
                let value = walk_expression(inner)?;
                args.push(Argument::positional(value));
            }
        }
    }
    Ok(args)
}

// ════════════════════════════════════════════════════════════════════════════
// String literal helpers
// ════════════════════════════════════════════════════════════════════════════

fn walk_string_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::raw_string => {
            let s = pair.as_str();
            let inner = if s.starts_with("r'") {
                &s[2..s.len() - 1]
            } else {
                &s[2..s.len() - 1]
            };
            Ok(ExprKind::Lit(Literal::Str(inner.to_string())))
        }
        Rule::interpolated_double_string | Rule::interpolated_single_string => {
            walk_interpolated_string(pair)
        }
        Rule::triple_double_string | Rule::triple_single_string => walk_interpolated_string(pair),
        _ => {
            // Fallback
            Ok(ExprKind::Lit(Literal::Str(unquote_string_literal(&pair))))
        }
    }
}

fn walk_interpolated_string(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut has_interp = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            // Opening/closing quotes
            Rule::dq_open
            | Rule::dq_close
            | Rule::sq_open
            | Rule::sq_close
            | Rule::triple_double_head
            | Rule::triple_double_tail
            | Rule::triple_single_head
            | Rule::triple_single_tail => {}

            // Text chunks
            Rule::dq_chars
            | Rule::sq_chars
            | Rule::triple_double_chars
            | Rule::triple_single_chars => {
                let text = unescape_string_chars(p.as_str());
                parts.push(InterpolPart::Text(text));
            }

            // Interpolation
            Rule::dq_interp
            | Rule::sq_interp
            | Rule::triple_double_interp
            | Rule::triple_single_interp => {
                has_interp = true;
                let inner = p.into_inner().next().ok_or("empty interpolation")?;
                match inner.as_rule() {
                    Rule::interp_simple => {
                        // $ident
                        let ident = inner.into_inner().next().ok_or("interp_simple: no ident")?;
                        parts.push(InterpolPart::Expr(Expression::ident(ident.as_str())));
                    }
                    Rule::interp_complex => {
                        // ${expr}
                        let expr_pair =
                            inner.into_inner().next().ok_or("interp_complex: no expr")?;
                        parts.push(InterpolPart::Expr(walk_expression(expr_pair)?));
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }

    if !has_interp {
        // Plain string with no interpolation
        let text: String = parts
            .iter()
            .map(|p| match p {
                InterpolPart::Text(s) => s.as_str(),
                _ => "",
            })
            .collect();
        Ok(ExprKind::Lit(Literal::Str(text)))
    } else {
        Ok(ExprKind::Interpolation(parts))
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let start = pair.as_span().start_pos().line_col();
    let end = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32 - 1,
        start_col: start.1 as u32 - 1,
        end_line: end.0 as u32 - 1,
        end_col: end.1 as u32 - 1,
    }
}

fn is_kw(r: Rule) -> bool {
    matches!(
        r,
        Rule::abstract_kw
            | Rule::as_kw
            | Rule::assert_kw
            | Rule::async_kw
            | Rule::await_kw
            | Rule::break_kw
            | Rule::case_kw
            | Rule::catch_kw
            | Rule::class_kw
            | Rule::const_kw
            | Rule::continue_kw
            | Rule::covariant_kw
            | Rule::default_kw
            | Rule::deferred_kw
            | Rule::do_kw
            | Rule::dynamic_kw
            | Rule::else_kw
            | Rule::enum_kw
            | Rule::export_kw
            | Rule::extends_kw
            | Rule::extension_kw
            | Rule::external_kw
            | Rule::factory_kw
            | Rule::false_kw
            | Rule::final_kw
            | Rule::finally_kw
            | Rule::for_kw
            | Rule::function_kw
            | Rule::hide_kw
            | Rule::if_kw
            | Rule::implements_kw
            | Rule::import_kw
            | Rule::in_kw
            | Rule::interface_kw
            | Rule::is_kw
            | Rule::late_kw
            | Rule::library_kw
            | Rule::mixin_kw
            | Rule::native_kw
            | Rule::new_kw
            | Rule::null_kw
            | Rule::on_kw
            | Rule::operator_kw
            | Rule::override_kw
            | Rule::part_kw
            | Rule::required_kw
            | Rule::rethrow_kw
            | Rule::return_kw
            | Rule::show_kw
            | Rule::static_kw
            | Rule::super_kw
            | Rule::switch_kw
            | Rule::sync_kw
            | Rule::this_kw
            | Rule::throw_kw
            | Rule::true_kw
            | Rule::try_kw
            | Rule::typedef_kw
            | Rule::var_keyword
            | Rule::void_kw
            | Rule::when_kw
            | Rule::while_kw
            | Rule::with_kw
            | Rule::yield_kw
    )
}

fn extract_type_name(pair: &Pair<Rule>) -> String {
    // Extract the base type name from a type_annotation, stripping generics and nullable
    let s = pair.as_str().trim();
    // For simple display, strip any <...> and ?
    if let Some(idx) = s.find('<') {
        s[..idx].trim().to_string()
    } else if s.ends_with('?') {
        s[..s.len() - 1].trim().to_string()
    } else {
        s.to_string()
    }
}

fn extract_type_name_from_clause(pair: &Pair<Rule>) -> Option<String> {
    for p in pair.clone().into_inner() {
        if p.as_rule() == Rule::type_annotation {
            return Some(extract_type_name(&p));
        }
    }
    None
}

fn extract_type_from_inner(pair: Pair<Rule>) -> String {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::type_annotation {
            return extract_type_name(&p);
        }
    }
    "dynamic".to_string()
}

fn unquote_string_literal(pair: &Pair<Rule>) -> String {
    let s = pair.as_str();
    // Handle raw strings
    if s.starts_with("r'") || s.starts_with("r\"") {
        return s[2..s.len() - 1].to_string();
    }
    // Handle triple-quoted strings
    if s.starts_with("'''") || s.starts_with("\"\"\"") {
        return unescape_string_chars(&s[3..s.len() - 3]);
    }
    // Handle single/double quoted
    if s.len() >= 2 {
        return unescape_string_chars(&s[1..s.len() - 1]);
    }
    s.to_string()
}

fn unescape_string_chars(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('n') => result.push('\n'),
                Some('t') => result.push('\t'),
                Some('r') => result.push('\r'),
                Some('\\') => result.push('\\'),
                Some('\'') => result.push('\''),
                Some('"') => result.push('"'),
                Some('$') => result.push('$'),
                Some('0') => result.push('\0'),
                Some('x') => {
                    let hi = chars.next();
                    let lo = chars.next();
                    if let (Some(hi), Some(lo)) = (hi, lo) {
                        let hex = format!("{}{}", hi, lo);
                        if let Ok(value) = u32::from_str_radix(&hex, 16) {
                            if let Some(ch) = char::from_u32(value) {
                                result.push(ch);
                            }
                        }
                    }
                }
                Some('u') => {
                    let mut hex = String::new();
                    for _ in 0..4 {
                        if let Some(ch) = chars.next() {
                            hex.push(ch);
                        }
                    }
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            result.push(ch);
                        }
                    }
                }
                Some(other) => {
                    result.push('\\');
                    result.push(other);
                }
                None => result.push('\\'),
            }
        } else {
            result.push(c);
        }
    }
    result
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::IDiv => BinOp::IDiv,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::Concat => BinOp::Concat,
    }
}
