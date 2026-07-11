//! Java walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! is shared with every other vybex language.
//!
//! ## Walker normalisations
//!
//! - **`System.out.println` / `System.out.print`**: rewritten in the walker
//!   to bare `println` / `print` calls so the profile can bind them.
//! - **`System.exit(code)`**: rewritten to `__process_exit(code)`.
//! - **Enhanced-for (`for (T x : iterable)`)**: `ForIn { of: true }` — Java
//!   always iterates values.
//! - **Implicit `super()` in child-class constructors**: injected when
//!   no explicit `super(...)` is found at the top of the ctor body.
//! - **Generic type arguments**: erased (Vybe is dynamic).
//! - **Char literals**: lowered to integer code-point literals.
//! - **Text blocks** (Java 15+): normalised to plain string literals.
//! - **Lambda params**: both typed `(T x) -> body` and untyped `x -> body`
//!   forms reduced to bare name params.

use super::{JavaParser, Rule};
use crate::ast::*;
use pest::iterators::Pair;
use pest::Parser;
use std::collections::{HashMap, HashSet};

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs =
        JavaParser::parse(Rule::program, source).map_err(|e| format!("Java parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI => continue,
            Rule::package_declaration => {}
            Rule::import_declaration => {
                if let Some(imp) = walk_import(&pair) {
                    imports.push(imp);
                }
            }
            Rule::class_declaration => body.push(Statement::new(walk_class(pair)?)),
            Rule::interface_declaration => body.push(Statement::new(walk_interface(pair)?)),
            Rule::enum_declaration => body.push(Statement::new(walk_enum_decl(pair)?)),
            Rule::record_declaration => body.push(Statement::new(walk_record(pair)?)),
            Rule::annotation_type_declaration => {}
            _ => {
                if let Some(s) = walk_statement(pair)? {
                    body.push(s);
                }
            }
        }
    }

    hoist_java_nested_types(&mut body);
    rewrite_java_user_tostring_calls(&mut body);
    normalize_java_class_tree(&mut body);

    // Java: inject a top-level call to the class's static main method.
    // Uses the same pattern as EntryPoint::Method in bundle.rs.
    if let Some(class_name) = find_main_class(&body) {
        body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class_name)),
                    field: "main".to_string(),
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            },
        ))));
    }

    Ok(Module {
        name: String::new(),
        language: Lang::Java,
        body,
        imports,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Span / helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let (line, col) = pair.line_col();
    Span {
        start_line: line as u32,
        start_col: col as u32,
        end_line: line as u32,
        end_col: col as u32,
    }
}

fn is_kw(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::final_kw
            | Rule::static_kw
            | Rule::public_kw
            | Rule::private_kw
            | Rule::protected_kw
            | Rule::abstract_kw
            | Rule::synchronized_kw
            | Rule::native_kw
            | Rule::transient_kw
            | Rule::volatile_kw
            | Rule::strictfp_kw
            | Rule::default_kw
            | Rule::sealed_kw
            | Rule::non_sealed_kw
            | Rule::var_kw
    )
}

// ════════════════════════════════════════════════════════════════════════════
// Imports
// ════════════════════════════════════════════════════════════════════════════

fn walk_import(pair: &Pair<Rule>) -> Option<Import> {
    let span = to_span(pair);
    let src = pair.as_str();
    // Skip star imports
    if src.contains(".*") {
        return None;
    }
    let text = src
        .trim_start_matches("import")
        .trim_start_matches(" static")
        .trim()
        .trim_end_matches(';')
        .trim();
    if text.is_empty() {
        return None;
    }
    let name = text.rsplit('.').next().unwrap_or(text).to_string();
    Some(Import {
        kind: ImportKind::Named {
            path: format!("java:{}", text.replace('.', "/")),
            names: vec![ImportName { name, alias: None }],
            level: 0,
        },
        span,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Modifiers
// ════════════════════════════════════════════════════════════════════════════

struct ParsedModifiers {
    visibility: Visibility,
    is_static: bool,
    is_abstract: bool,
}

impl Default for ParsedModifiers {
    fn default() -> Self {
        Self {
            visibility: Visibility::Public,
            is_static: false,
            is_abstract: false,
        }
    }
}

fn parse_modifiers(pair: &Pair<Rule>) -> ParsedModifiers {
    let mut out = ParsedModifiers::default();
    if pair.as_rule() != Rule::modifiers {
        return out;
    }
    for m in pair.clone().into_inner() {
        if m.as_rule() == Rule::modifier {
            let inner = m.into_inner().next();
            match inner.as_ref().map(|p| p.as_rule()) {
                Some(Rule::private_kw) => out.visibility = Visibility::Private,
                Some(Rule::protected_kw) => out.visibility = Visibility::Protected,
                Some(Rule::public_kw) => out.visibility = Visibility::Public,
                Some(Rule::static_kw) => out.is_static = true,
                Some(Rule::abstract_kw) => out.is_abstract = true,
                _ => {}
            }
        }
    }
    out
}

fn into_modifiers(pm: ParsedModifiers) -> Modifiers {
    Modifiers {
        visibility: pm.visibility,
        is_static: pm.is_static,
        is_abstract: pm.is_abstract,
        ..Modifiers::default()
    }
}

fn into_class_modifiers(pm: ParsedModifiers) -> ClassModifiers {
    ClassModifiers {
        visibility: pm.visibility,
        is_abstract: pm.is_abstract,
        is_static: pm.is_static,
        ..ClassModifiers::default()
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Class
// ════════════════════════════════════════════════════════════════════════════

fn walk_class(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut class_modifiers = ClassModifiers::default();

    let mut inner = pair.into_inner().peekable();

    // modifiers
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        let mp = inner.next().unwrap();
        class_modifiers = into_class_modifiers(parse_modifiers(&mp));
    }
    // "class" keyword matched as ident_name in grammar — next ident is the name
    if let Some(n) = inner.next() {
        name = n.as_str().to_string();
    }

    for p in inner {
        match p.as_rule() {
            Rule::type_params => {}
            Rule::type_ref => {
                // extends clause: first type_ref
                if parents.is_empty() {
                    parents.push(extract_ref_name(&p));
                }
            }
            Rule::type_ref_list => {
                for tr in p.into_inner() {
                    if tr.as_rule() == Rule::type_ref {
                        interfaces.push(extract_ref_name(&tr));
                    }
                }
            }
            Rule::class_body => {
                members = walk_class_body(p)?;
            }
            _ => {}
        }
    }

    if !parents.is_empty() {
        inject_implicit_super(&mut members);
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: class_modifiers,
        decorators: vec![],
    })
}

fn walk_class_body(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constructor_declaration => members.push(walk_constructor(p)?),
            Rule::method_declaration | Rule::default_method_declaration => {
                members.push(walk_method(p)?)
            }
            Rule::field_declaration => {
                for m in walk_field(p)? {
                    members.push(m);
                }
            }
            Rule::class_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_class(p)?,
                ))));
            }
            Rule::interface_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_interface(p)?,
                ))));
            }
            Rule::enum_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_enum_decl(p)?,
                ))));
            }
            Rule::record_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_record(p)?,
                ))));
            }
            Rule::static_initializer | Rule::instance_initializer => {
                // treat as method named __init_block__
                let body: Vec<Statement> = p
                    .into_inner()
                    .filter_map(|b| {
                        if b.as_rule() == Rule::block_statement {
                            walk_block(b).ok()
                        } else {
                            None
                        }
                    })
                    .flatten()
                    .collect();
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: "__init_block__".to_string(),
                        params: vec![],
                        return_type: None,
                        body,
                        modifiers: Modifiers::default(),
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    },
                ))));
            }
            Rule::annotation => {}
            _ => {}
        }
    }
    Ok(members)
}

fn walk_constructor(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let visibility = pm.visibility;

    // constructor name — same as class, skip
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::ident_name) {
        inner.next();
    }
    // skip optional type_params
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        inner.next();
    }

    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut base_args: Option<Vec<Expression>> = None;
    let mut initializer_target = ConstructorInitializerTarget::Base;

    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::throws_clause => {}
            Rule::function_body_block => {
                body = walk_block(p)?;
                // Extract super(...) or this(...) call from top of body
                extract_base_call_from_body(&mut body, &mut base_args, &mut initializer_target);
            }
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        params,
        body,
        base_args,
        initializer_target,
        visibility,
    })
}

fn walk_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let modifiers = into_modifiers(pm);

    // Return type (type_ref)
    let return_type = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        let tr = inner.next().unwrap();
        Some(extract_ref_name(&tr))
    } else {
        None
    };

    let name = inner
        .next()
        .ok_or("method: missing name")?
        .as_str()
        .to_string();

    // skip type_params
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        inner.next();
    }

    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::dim_suffix => {}
            Rule::throws_clause => {}
            Rule::function_body => {
                for fb in p.into_inner() {
                    if fb.as_rule() == Rule::function_body_block {
                        body = walk_block(fb)?;
                    }
                }
            }
            _ => {}
        }
    }

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
    ))))
}

fn walk_field(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let modifiers = into_modifiers(pm);

    // type_ref
    let type_hint = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };

    let mut fields = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            let mut di = p.into_inner().peekable();
            let name = di.next().ok_or("field: missing name")?.as_str().to_string();
            // skip dim_suffix(s)
            while di.peek().map(|x| x.as_rule()) == Some(Rule::dim_suffix) {
                di.next();
            }
            let init = if di.peek().map(|x| x.as_rule()) == Some(Rule::initializer) {
                Some(walk_initializer(di.next().unwrap())?)
            } else {
                type_hint.as_deref().and_then(default_expr_for_java_type)
            };
            fields.push(ClassMember::Field {
                name,
                type_hint: type_hint.clone(),
                init,
                modifiers: modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
        }
    }
    Ok(fields)
}

// ════════════════════════════════════════════════════════════════════════════
// Interface
// ════════════════════════════════════════════════════════════════════════════

fn walk_interface(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("interface: missing name")?
        .as_str()
        .to_string();
    // skip type_params
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        inner.next();
    }

    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_ref_list => {
                for tr in p.into_inner() {
                    if tr.as_rule() == Rule::type_ref {
                        parents.push(extract_ref_name(&tr));
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::method_declaration | Rule::default_method_declaration => {
                            members.push(walk_method(m)?);
                        }
                        Rule::field_declaration => {
                            for f in walk_field(m)? {
                                members.push(f);
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
        interfaces: vec![],
        members,
        modifiers: into_class_modifiers(pm),
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Enum
// ════════════════════════════════════════════════════════════════════════════

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("enum: missing name")?
        .as_str()
        .to_string();

    let mut enum_members: Vec<EnumMember> = Vec::new();

    for p in inner {
        if p.as_rule() == Rule::enum_values {
            for ev in p.into_inner() {
                if ev.as_rule() == Rule::enum_value {
                    let val_name = ev
                        .into_inner()
                        .find(|x| x.as_rule() == Rule::ident_name)
                        .map(|x| x.as_str().to_string())
                        .unwrap_or_default();
                    if !val_name.is_empty() {
                        enum_members.push(EnumMember {
                            name: val_name,
                            value: None,
                            constructor_args: vec![],
                        });
                    }
                }
            }
        }
    }

    Ok(StmtKind::EnumDecl {
        name,
        members: enum_members,
        visibility: pm.visibility,
        is_flags: false,
        backing_type: None,
        interfaces: vec![],
        body_members: vec![],
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Record
// ════════════════════════════════════════════════════════════════════════════

fn walk_record(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("record: missing name")?
        .as_str()
        .to_string();
    // skip type_params
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        inner.next();
    }

    let mut component_params: Vec<Param> = Vec::new();
    let mut extra_members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::record_component_list => {
                for comp in p.into_inner() {
                    if comp.as_rule() == Rule::record_component {
                        let mut ci = comp.into_inner().peekable();
                        // skip annotations
                        while ci.peek().map(|x| x.as_rule()) == Some(Rule::annotation) {
                            ci.next();
                        }
                        // skip type_ref
                        if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                            ci.next();
                        }
                        if let Some(n) = ci.next() {
                            component_params.push(Param {
                                name: n.as_str().to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                        }
                    }
                }
            }
            Rule::type_ref_list => {}
            Rule::class_body => {
                extra_members = walk_class_body(p)?;
            }
            _ => {}
        }
    }

    // Synthesise a constructor from the record components
    let ctor_body: Vec<Statement> = component_params
        .iter()
        .map(|p| {
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: p.name.clone(),
                    null_safe: false,
                })],
                value: Expression::new(ExprKind::Ident(p.name.clone())),
            })
        })
        .collect();

    let mut members = vec![ClassMember::Constructor {
        params: component_params,
        body: ctor_body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }];
    members.extend(extra_members);

    Ok(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces: vec![],
        members,
        modifiers: into_class_modifiers(pm),
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => return Ok(None),

        Rule::block_statement => StmtKind::Block(walk_block(pair)?),

        Rule::variable_declaration_statement => walk_var_decl(pair)?,

        Rule::if_statement => walk_if(pair)?,

        Rule::for_statement => walk_for_stmt(pair)?,

        Rule::enhanced_for_statement => walk_enhanced_for(pair)?,

        Rule::while_statement => {
            let mut inner = pair.into_inner();
            let cond = walk_expr_inner(&mut inner)?;
            let body = walk_body_inner(&mut inner)?;
            StmtKind::While {
                cond,
                body,
                else_body: None,
            }
        }

        Rule::do_while_statement => {
            let mut inner = pair.into_inner();
            let body_pair = inner.next().ok_or("do-while: missing body")?;
            let body = walk_statement_into_body(body_pair)?;
            let cond = walk_expr_inner(&mut inner)?;
            StmtKind::DoWhile {
                body,
                cond,
                until: false,
            }
        }

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let e = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(e)
        }

        Rule::break_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Break(match label {
                Some(l) if !l.is_empty() => BreakTarget::Label(l),
                _ => BreakTarget::Implicit,
            })
        }

        Rule::continue_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Continue(match label {
                Some(l) if !l.is_empty() => ContinueTarget::Label(l),
                _ => ContinueTarget::Implicit,
            })
        }

        Rule::throw_statement => {
            let inner = pair.into_inner().next().ok_or("throw: missing expr")?;
            StmtKind::Throw {
                expr: Some(walk_expression(inner)?),
                cause: None,
            }
        }

        Rule::try_statement | Rule::try_with_resources_statement => walk_try(pair)?,

        Rule::assert_statement => {
            let mut exprs: Vec<Expression> = Vec::new();
            for p in pair.into_inner() {
                if !is_kw(p.as_rule()) {
                    exprs.push(walk_expression(p)?);
                }
            }
            let test = exprs.remove(0);
            let msg = exprs.into_iter().next();
            StmtKind::Assert { test, msg }
        }

        Rule::yield_statement => {
            let e = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(e)
        }

        Rule::labeled_statement => {
            // Strip label; walk inner statement
            let inner = pair
                .into_inner()
                .find(|p| !matches!(p.as_rule(), Rule::ident_name));
            if let Some(s) = inner {
                return walk_statement(s);
            }
            return Ok(None);
        }

        Rule::synchronized_statement => {
            // synchronized (lock) { body } → just compile the body block
            let mut inner = pair.into_inner();
            let _lock = walk_expr_inner(&mut inner)?;
            let body_pair = inner.next().ok_or("synchronized: missing body")?;
            StmtKind::Block(walk_block(body_pair)?)
        }

        Rule::super_constructor_call => {
            let args: Vec<Argument> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::argument_list)
                .flat_map(|al| walk_arguments(al).unwrap_or_default())
                .collect();
            StmtKind::Expr(Expression::new(ExprKind::SuperCall { method: None, args }))
        }

        Rule::this_constructor_call => {
            let args: Vec<Argument> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::argument_list)
                .flat_map(|al| walk_arguments(al).unwrap_or_default())
                .collect();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::This)),
                args,
                optional: false,
            }))
        }

        Rule::expression_statement => {
            let inner = pair.into_inner().next().ok_or("expr stmt: missing expr")?;
            StmtKind::Expr(walk_expression(inner)?)
        }

        Rule::local_class_declaration => {
            let cls = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::class_declaration);
            if let Some(c) = cls {
                walk_class(c)?
            } else {
                return Ok(None);
            }
        }

        Rule::local_record_declaration => {
            let rec = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::record_declaration);
            if let Some(r) = rec {
                walk_record(r)?
            } else {
                return Ok(None);
            }
        }

        Rule::class_declaration => walk_class(pair)?,
        Rule::interface_declaration => walk_interface(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::record_declaration => walk_record(pair)?,

        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if let Some(s) = walk_statement(p)? {
            out.push(s);
        }
    }
    Ok(out)
}

fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if pair.as_rule() == Rule::block_statement {
        walk_block(pair)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(vec![]),
        }
    }
}

/// Pull the next `expression`-shaped child from `inner` and walk it.
fn walk_expr_inner<'a>(
    inner: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Expression, String> {
    walk_expression(inner.next().ok_or("missing expression")?)
}

/// Pull the next statement-shaped child from `inner` and expand to body.
fn walk_body_inner<'a>(
    inner: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Vec<Statement>, String> {
    let p = inner.next().ok_or("missing body")?;
    walk_statement_into_body(p)
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();

    let is_final = if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
        true
    } else {
        false
    };

    let kind = if is_final {
        VarDeclKind::Const
    } else {
        VarDeclKind::Let
    };

    // var_kw or type_ref
    let type_hint: Option<String> = if inner.peek().map(|p| p.as_rule()) == Some(Rule::var_kw) {
        inner.next();
        None
    } else if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };

    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            declarations.push(walk_var_declarator(p, type_hint.clone())?);
        }
    }

    Ok(StmtKind::VarDecl { declarations, kind })
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
            Rule::dim_suffix => {}
            Rule::initializer => init = Some(walk_initializer(p)?),
            _ => {}
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

fn walk_initializer(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("initializer: empty")?;
    match inner.as_rule() {
        Rule::array_initializer => {
            let mut elems = Vec::new();
            for el in inner.into_inner() {
                if el.as_rule() == Rule::initializer {
                    elems.push(ArrayElement {
                        key: None,
                        value: walk_initializer(el)?,
                        spread: false,
                        by_ref: false,
                    });
                }
            }
            Ok(Expression::new(ExprKind::Array(elems)))
        }
        _ => walk_expression(inner),
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expr_inner(&mut inner)?;
    let then_pair = inner.next().ok_or("if: missing then")?;
    let then_body = walk_statement_into_body(then_pair)?;

    let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;

    if let Some(else_pair) = inner.next() {
        // peek inside — if it's an if_statement, it's an else-if
        if else_pair.as_rule() == Rule::if_statement {
            if let StmtKind::If {
                cond: elif_cond,
                then_body: elif_body,
                elifs: nested_elifs,
                else_body: nested_else,
            } = walk_if(else_pair)?
            {
                elifs.push((elif_cond, elif_body));
                elifs.extend(nested_elifs);
                else_body = nested_else;
            }
        } else {
            else_body = Some(walk_statement_into_body(else_pair)?);
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_for_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner().peekable();

    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body: Vec<Statement> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::for_init => init = Some(Box::new(walk_for_init(p)?)),
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(p)?);
                } else {
                    // update expression
                    update = Some(walk_expression(p)?);
                }
            }
            Rule::for_update => {
                // for_update = { expression ~ ("," ~ expression)* }
                let mut exprs: Vec<Expression> = Vec::new();
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::expression {
                        exprs.push(walk_expression(ep)?);
                    }
                }
                if exprs.len() == 1 {
                    update = Some(exprs.remove(0));
                } else if exprs.len() > 1 {
                    update = Some(Expression::new(ExprKind::Sequence(exprs)));
                }
            }
            _ => {
                body = walk_statement_into_body(p)?;
            }
        }
    }

    Ok(StmtKind::For {
        init,
        cond,
        update,
        body,
    })
}

fn walk_for_init(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner().peekable();

    // optional final
    let is_final = if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
        true
    } else {
        false
    };

    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        // variable declaration
        let type_hint = Some(extract_ref_name(&inner.next().unwrap()));
        let kind = if is_final {
            VarDeclKind::Const
        } else {
            VarDeclKind::Let
        };
        let mut decls = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::var_declarator {
                decls.push(walk_var_declarator(p, type_hint.clone())?);
            }
        }
        return Ok(Statement::new(StmtKind::VarDecl {
            declarations: decls,
            kind,
        }));
    }

    // expression list
    let mut exprs: Vec<Expression> = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::expression {
            exprs.push(walk_expression(p)?);
        }
    }
    if exprs.len() == 1 {
        let e = exprs.remove(0);
        Ok(Statement::new(StmtKind::Expr(e)))
    } else {
        Ok(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Sequence(exprs),
        ))))
    }
}

fn walk_enhanced_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();

    if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
    }
    // type_ref or var_kw — skip either
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref)
        || inner.peek().map(|p| p.as_rule()) == Some(Rule::var_kw)
    {
        inner.next();
    }

    let var = inner
        .next()
        .ok_or("for-each: missing var")?
        .as_str()
        .to_string();
    let iter = walk_expr_inner(&mut inner)?;
    let body = walk_body_inner(&mut inner)?;

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter,
        body,
        of: true,
        else_body: None,
        is_async: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let switch_expr = walk_expr_inner(&mut inner)?;
    let value_name = format!("__java_switch_value_{}_{}", span.start_line, span.start_col);
    let matched_name = format!(
        "__java_switch_matched_{}_{}",
        span.start_line, span.start_col
    );
    let done_name = format!("__java_switch_done_{}_{}", span.start_line, span.start_col);

    let mut arms: Vec<JavaSwitchArm> = Vec::new();
    let mut all_label_matches: Vec<Expression> = Vec::new();

    for case_pair in inner {
        if case_pair.as_rule() != Rule::switch_case {
            continue;
        }
        let mut ci = case_pair.into_inner().peekable();
        let mut labels: Vec<Expression> = Vec::new();
        let mut body: Vec<Statement> = Vec::new();
        let mut is_default = false;
        let mut is_arrow = false;
        let src = {
            let tmp = ci
                .peek()
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            tmp
        };

        if src.trim() == "default" {
            is_default = true;
            ci.next(); // consume "default"
        }

        for p in ci {
            match p.as_rule() {
                Rule::switch_label => {
                    if let Ok(e) = walk_switch_label(p) {
                        let label = java_switch_label_expr(e);
                        all_label_matches.push(java_switch_label_match(&value_name, label.clone()));
                        labels.push(label);
                    }
                }
                Rule::switch_rule_body => {
                    is_arrow = true;
                    for rb in p.into_inner() {
                        body.extend(walk_switch_rule_body_part(rb)?);
                    }
                }
                _ => {
                    if let Some(s) = walk_statement(p)? {
                        body.push(s);
                    }
                }
            }
        }

        let (body, has_break) = java_strip_top_level_switch_break(body);
        let is_default_arm = is_default || labels.is_empty();
        arms.push(JavaSwitchArm {
            labels,
            body,
            is_default: is_default_arm,
            has_break: has_break || is_arrow,
        });
    }

    let any_label_match =
        java_or_exprs(all_label_matches).unwrap_or_else(|| Expression::bool(false));
    let mut lowered = vec![
        java_var_decl(&value_name, Some(switch_expr)),
        java_var_decl(&matched_name, Some(Expression::bool(false))),
        java_var_decl(&done_name, Some(Expression::bool(false))),
    ];

    for arm in arms {
        let raw_cond = if arm.is_default {
            java_binary(
                BinOp::Or,
                Expression::ident(&matched_name),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(any_label_match.clone()),
                }),
            )
        } else {
            let label_cond = java_or_exprs(
                arm.labels
                    .into_iter()
                    .map(|label| java_switch_label_match(&value_name, label))
                    .collect(),
            )
            .unwrap_or_else(|| Expression::bool(false));
            java_binary(BinOp::Or, Expression::ident(&matched_name), label_cond)
        };
        let cond = java_binary(
            BinOp::And,
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(Expression::ident(&done_name)),
            }),
            raw_cond,
        );
        let mut then_body = vec![java_assign_stmt(&matched_name, Expression::bool(true))];
        then_body.extend(arm.body);
        if arm.has_break {
            then_body.push(java_assign_stmt(&done_name, Expression::bool(true)));
            then_body.push(java_assign_stmt(&matched_name, Expression::bool(false)));
        }
        lowered.push(Statement::new(StmtKind::If {
            cond,
            then_body,
            elifs: vec![],
            else_body: None,
        }));
    }

    Ok(StmtKind::Block(lowered))
}

struct JavaSwitchArm {
    labels: Vec<Expression>,
    body: Vec<Statement>,
    is_default: bool,
    has_break: bool,
}

fn walk_switch_rule_body_part(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    match pair.as_rule() {
        Rule::block_statement => walk_block(pair),
        Rule::throw_statement => Ok(vec![Statement::new(walk_statement(pair)?.unwrap().kind)]),
        Rule::expression => Ok(vec![Statement::new(StmtKind::Expr(walk_expression(pair)?))]),
        Rule::expression_statement => Ok(walk_statement(pair)?.into_iter().collect()),
        _ => Ok(walk_statement(pair)?.into_iter().collect()),
    }
}

fn java_strip_top_level_switch_break(body: Vec<Statement>) -> (Vec<Statement>, bool) {
    let mut out = Vec::new();
    for stmt in body {
        if matches!(stmt.kind, StmtKind::Break(BreakTarget::Implicit)) {
            return (out, true);
        }
        out.push(stmt);
    }
    (out, false)
}

fn java_switch_label_expr(expr: Expression) -> Expression {
    expr
}

fn walk_switch_label(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut is_negative = false;
    let mut value = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::unary_op if p.as_str() == "-" => is_negative = true,
            Rule::literal => value = Some(walk_literal(p)?),
            Rule::qualified_name => {
                let text = p.as_str();
                if text.contains('.') {
                    let mut parts = text.split('.');
                    let first = parts.next().unwrap_or_default();
                    let mut expr = Expression::ident(first);
                    for part in parts {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: part.to_string(),
                            null_safe: false,
                        });
                    }
                    value = Some(expr);
                } else {
                    value = Some(Expression::ident(text));
                }
            }
            _ => {}
        }
    }
    let expr = value.unwrap_or_else(Expression::null);
    if is_negative {
        Ok(Expression::new(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: Box::new(expr),
        }))
    } else {
        Ok(expr)
    }
}

fn java_switch_label_match(value_name: &str, label: Expression) -> Expression {
    java_binary(BinOp::Eq, Expression::ident(value_name), label)
}

fn java_or_exprs(mut exprs: Vec<Expression>) -> Option<Expression> {
    let first = exprs.pop()?;
    Some(
        exprs
            .into_iter()
            .fold(first, |acc, expr| java_binary(BinOp::Or, expr, acc)),
    )
}

fn java_var_decl(name: &str, init: Option<Expression>) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init,
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn java_assign_stmt(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(name)],
        value,
    })
}

fn walk_switch_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let subject = walk_expr_inner(&mut inner)?;
    let mut arms: Vec<(Vec<Expression>, Expression)> = Vec::new();
    let mut default_expr: Option<Expression> = None;

    for arm in inner {
        if arm.as_rule() != Rule::switch_expr_arm {
            continue;
        }
        let mut labels = Vec::new();
        let mut value = None;
        let mut is_default = false;
        let mut ai = arm.into_inner().peekable();
        let src = ai
            .peek()
            .map(|p| p.as_str().to_string())
            .unwrap_or_default();
        if src.trim() == "default" {
            is_default = true;
            ai.next();
        }
        for p in ai {
            match p.as_rule() {
                Rule::switch_label => {
                    labels.push(java_switch_label_expr(walk_switch_label(p)?));
                }
                Rule::switch_rule_body => {
                    value = java_switch_rule_body_expr(p)?;
                }
                _ => {}
            }
        }
        if let Some(expr) = value {
            if is_default || labels.is_empty() {
                default_expr = Some(expr);
            } else {
                arms.push((labels, expr));
            }
        }
    }

    let mut result = default_expr.unwrap_or_else(Expression::null);
    for (labels, value) in arms.into_iter().rev() {
        let cond = java_or_exprs(
            labels
                .into_iter()
                .map(|label| java_binary(BinOp::Eq, subject.clone(), label))
                .collect(),
        )
        .unwrap_or_else(|| Expression::bool(false));
        result = java_ternary(cond, value, result);
    }
    Ok(result)
}

fn java_switch_rule_body_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => return Ok(Some(walk_expression(p)?)),
            Rule::expression_statement => {
                if let Some(expr_p) = p.into_inner().next() {
                    return Ok(Some(walk_expression(expr_p)?));
                }
            }
            Rule::block_statement => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::yield_statement {
                        let expr = stmt_pair
                            .into_inner()
                            .find(|inner| !is_kw(inner.as_rule()))
                            .map(walk_expression)
                            .transpose()?;
                        return Ok(expr);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body: Vec<Statement> = Vec::new();
    let mut catches: Vec<CatchClause> = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::resource => {}
            Rule::block_statement | Rule::function_body_block => {
                body = walk_block(p)?;
            }
            Rule::catch_clause => {
                let mut ci = p.into_inner().peekable();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::final_kw) {
                    ci.next();
                }
                let mut types: Vec<String> = Vec::new();
                while ci.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                    types.push(extract_ref_name(&ci.next().unwrap()));
                }
                let var_name = ci.next().map(|p| p.as_str().to_string());
                let catch_body = ci
                    .next()
                    .map(|b| walk_block(b))
                    .transpose()?
                    .unwrap_or_default();
                if types.is_empty() {
                    types.push("Exception".to_string());
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                if let Some(blk) = p.into_inner().next() {
                    finally = Some(walk_block(blk)?);
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

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param {
            out.push(walk_param(p)?);
        }
    }
    Ok(out)
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut inner = pair.into_inner().peekable();

    // annotations
    while inner.peek().map(|p| p.as_rule()) == Some(Rule::annotation) {
        inner.next();
    }
    // optional final
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
    }
    // type_ref
    let type_hint = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };
    // varargs
    let is_rest = if inner.peek().map(|p| p.as_rule()) == Some(Rule::varargs_marker) {
        inner.next();
        true
    } else {
        false
    };
    let name = inner
        .next()
        .ok_or("param: missing name")?
        .as_str()
        .to_string();
    // skip dim_suffix(s)
    while inner.peek().map(|p| p.as_rule()) == Some(Rule::dim_suffix) {
        inner.next();
    }

    Ok(Param {
        name,
        type_hint,
        default: None,
        pass_by: PassBy::Value,
        is_rest,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::expression => {
            let mut parts: Vec<Expression> = pair
                .into_inner()
                .map(walk_expression)
                .collect::<Result<_, _>>()?;
            if parts.len() == 1 {
                Ok(parts.remove(0))
            } else {
                Ok(Expression::new(ExprKind::Sequence(parts)))
            }
        }
        Rule::assignment_expression => walk_assignment(pair),
        Rule::ternary_expression => walk_ternary(pair),
        Rule::binop_expression => walk_binop(pair),
        Rule::instanceof_expression => walk_instanceof(pair),
        Rule::unary_expression => walk_unary(pair),
        Rule::cast_expression => {
            let mut ci = pair.into_inner();
            let cast_type = ci.next(); // cast_type
            if let (Some(cast_type), Some(operand)) = (cast_type, ci.next()) {
                let expr = walk_expression(operand)?;
                let ty = cast_type.as_str();
                if matches!(ty, "int" | "long" | "short" | "byte" | "char") {
                    Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__java_trunc_cast")),
                        args: vec![Argument::positional(expr)],
                        optional: false,
                    }))
                } else {
                    Ok(expr)
                }
            } else {
                Ok(Expression::null())
            }
        }
        Rule::postfix_expression => walk_postfix(pair),
        Rule::primary_chain => walk_primary_chain(pair),
        Rule::primary_atom => walk_primary_atom(pair),
        Rule::lambda_expression => walk_lambda(pair),
        Rule::switch_expression => walk_switch_expression(pair),
        _ => {
            let mut inner = pair.into_inner();
            if let Some(first) = inner.next() {
                walk_expression(first)
            } else {
                Ok(Expression::null())
            }
        }
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner().peekable();
    let first = inner.next().ok_or("assignment: empty")?;

    // Check if next is an assignment_op
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::assignment_op) {
        let op_str = inner.next().unwrap().as_str().to_string();
        let rhs = walk_expression(inner.next().ok_or("assignment: missing rhs")?)?;
        let lhs = walk_expression(first)?;

        if op_str == "=" {
            return Ok(Expression::new(ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(rhs),
            }));
        }
        // Compound assignment: `x += v` → `x = x + v`
        let bin_op = compound_op_to_binop(&op_str);
        return Ok(Expression::new(ExprKind::Assign {
            target: Box::new(lhs.clone()),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: bin_op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            })),
        }));
    }

    walk_expression(first)
}

fn walk_ternary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("ternary: missing cond")?)?;
    if let Some(then_p) = inner.next() {
        let then_e = walk_expression(then_p)?;
        let else_e = walk_expression(inner.next().ok_or("ternary: missing else")?)?;
        Ok(Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_e),
            else_: Box::new(else_e),
        }))
    } else {
        Ok(cond)
    }
}

fn walk_binop(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut left = walk_expression(inner.next().ok_or("binop: missing lhs")?)?;

    while let Some(op_pair) = inner.next() {
        let rhs = walk_expression(inner.next().ok_or("binop: missing rhs")?)?;
        let op = str_to_binop(op_pair.as_str().trim());
        if op == BinOp::Add
            && (is_java_string_concat_operand(&left) || is_java_string_concat_operand(&rhs))
        {
            left = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_concat")),
                args: vec![Argument::positional(left), Argument::positional(rhs)],
                optional: false,
            });
            continue;
        }
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(rhs),
        });
    }
    Ok(left)
}

fn is_java_string_concat_operand(expr: &Expression) -> bool {
    match expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Call { ref callee, .. } => {
            matches!(callee.kind, ExprKind::Ident(ref name) if name == "__java_string_concat")
        }
        _ => false,
    }
}

fn walk_instanceof(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut base = walk_expression(inner.next().ok_or("instanceof: empty")?)?;

    let mut saw_instanceof = false;
    for p in inner {
        match p.as_rule() {
            Rule::instanceof_kw => {
                saw_instanceof = true;
            }
            Rule::type_ref if saw_instanceof => {
                let type_name = extract_ref_name(&p);
                base = Expression::new(ExprKind::IsType {
                    expr: Box::new(base),
                    type_name,
                });
                saw_instanceof = false;
            }
            Rule::ident_name => {} // pattern binding var — skip
            _ => {}
        }
    }
    Ok(base)
}

fn walk_unary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("unary: empty")?;

    if first.as_rule() == Rule::unary_op {
        let op_str = first.as_str();
        let operand = walk_expression(inner.next().ok_or("unary: missing operand")?)?;
        let op = match op_str {
            "++" => UnaryOp::PreInc,
            "--" => UnaryOp::PreDec,
            "!" => UnaryOp::Not,
            "-" => UnaryOp::Neg,
            "+" => UnaryOp::Pos,
            "~" => UnaryOp::BitNot,
            _ => UnaryOp::Not,
        };
        return Ok(Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(operand),
        }));
    }
    walk_expression(first)
}

fn walk_postfix(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let base = walk_expression(inner.next().ok_or("postfix: empty")?)?;
    if let Some(op) = inner.next() {
        let unop = match op.as_str() {
            "++" => UnaryOp::PostInc,
            "--" => UnaryOp::PostDec,
            _ => UnaryOp::PostInc,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: unop,
            expr: Box::new(base),
        }))
    } else {
        Ok(base)
    }
}

fn walk_primary_chain(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut current = walk_expression(inner.next().ok_or("chain: empty")?)?;

    for chain in inner {
        // chain_suffix is a wrapper rule — unwrap to get the inner rule
        let chain = if chain.as_rule() == Rule::chain_suffix {
            chain.into_inner().next().unwrap_or_else(|| unreachable!())
        } else {
            chain
        };
        match chain.as_rule() {
            Rule::method_call_suffix => {
                let mut ci = chain.into_inner().peekable();
                let method_name = ci
                    .next()
                    .ok_or("method call: missing name")?
                    .as_str()
                    .to_string();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let args = if let Some(al) = ci.next() {
                    if al.as_rule() == Rule::argument_list {
                        walk_arguments(al)?
                    } else {
                        vec![]
                    }
                } else {
                    vec![]
                };
                current = normalise_method_call(current, method_name, args);
            }
            Rule::member_access_suffix => {
                let field = chain
                    .into_inner()
                    .next()
                    .ok_or("member: empty")?
                    .as_str()
                    .to_string();
                current = Expression::new(ExprKind::Member {
                    object: Box::new(current),
                    field,
                    null_safe: false,
                });
            }
            Rule::index_suffix => {
                let idx = walk_expression(chain.into_inner().next().ok_or("index: empty")?)?;
                current = Expression::new(ExprKind::Index {
                    object: Box::new(current),
                    index: Box::new(idx),
                    null_safe: false,
                });
            }
            Rule::call_suffix => {
                // Bare function call: callee(args) — the base is the callee.
                let mut ci = chain.into_inner().peekable();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let args = if let Some(al) = ci.next() {
                    if al.as_rule() == Rule::argument_list {
                        walk_arguments(al)?
                    } else {
                        vec![]
                    }
                } else {
                    vec![]
                };
                if let ExprKind::Member {
                    object,
                    field,
                    null_safe: _,
                } = current.kind
                {
                    current = normalise_method_call(*object, field, args);
                    continue;
                }
                current = Expression::new(ExprKind::Call {
                    callee: Box::new(current),
                    args,
                    optional: false,
                });
            }
            _ => {}
        }
    }
    Ok(current)
}

/// Normalise Java-specific call patterns to a compiler-friendly shape.
fn normalise_method_call(receiver: Expression, method: String, args: Vec<Argument>) -> Expression {
    // System.out.println(x) → println(x)
    // System.out.print(x)   → print(x)
    // receiver = Member { Ident("System"), "out" }, method = "println"
    if let ExprKind::Member {
        object: ref root_obj,
        field: ref root_field,
        ..
    } = receiver.kind
    {
        if let ExprKind::Ident(ref root_name) = root_obj.kind {
            if root_name == "System"
                && root_field == "out"
                && matches!(
                    method.as_str(),
                    "println" | "print" | "printf" | "format" | "append"
                )
            {
                let normalized = match method.as_str() {
                    "format" => "printf",
                    "append" => "print",
                    _ => &method,
                };
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(normalized)),
                    args,
                    optional: false,
                });
            }
            // System.exit(code) → __process_exit(code)
            if root_name == "System" && root_field == "exit" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__process_exit")),
                    args,
                    optional: false,
                });
            }
        }
    }

    if java_expr_dotted_name(&receiver).as_deref() == Some("java.math.BigInteger")
        && method == "valueOf"
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_bigint")),
            args,
            optional: false,
        });
    }

    // `List.remove(int)` and `List.remove(Object)` are distinct Java overloads.
    // The boxed form is explicit in the parsed tree, so preserve that distinction
    // before profile dispatch erases the receiver type.
    if method == "remove"
        && args.len() == 1
        && matches!(
            args[0].value.kind,
            ExprKind::Call { ref callee, .. }
                if matches!(callee.kind, ExprKind::Ident(ref name) if name == "Integer.valueOf")
        )
    {
        let mut call_args = Vec::with_capacity(2);
        call_args.push(Argument::positional(receiver));
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_list_remove_value")),
            args: call_args,
            optional: false,
        });
    }

    if let Some(type_name) = java_qualified_static_type(&receiver) {
        if type_name == "BigInteger" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args,
                optional: false,
            });
        }
        if type_name == "Comparator" {
            if let Some(expr) = normalise_comparator_static_call(&method, args.clone()) {
                return expr;
            }
        }
        if type_name == "String" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_value_of")),
                args,
                optional: false,
            });
        }
        if type_name == "String" && method == "format" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_format")),
                args,
                optional: false,
            });
        }
        if type_name == "String" && method == "join" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_join")),
                args,
                optional: false,
            });
        }
        let dotted = format!("{}.{}", type_name, method);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(&dotted)),
            args,
            optional: false,
        });
    }

    // Static type method calls: Integer.parseInt("42") → call "Integer.parseInt"
    // The profile has dotted builtins like "Integer.parseInt", "Math.max", etc.
    if let ExprKind::Ident(ref type_name) = receiver.kind {
        if is_java_type_or_util(type_name) {
            if type_name == "Comparator" {
                if let Some(expr) = normalise_comparator_static_call(&method, args.clone()) {
                    return expr;
                }
            }
            if type_name == "String" && method == "valueOf" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_value_of")),
                    args,
                    optional: false,
                });
            }
            if type_name == "String" && method == "format" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_format")),
                    args,
                    optional: false,
                });
            }
            if type_name == "String" && method == "join" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_join")),
                    args,
                    optional: false,
                });
            }
            let dotted = format!("{}.{}", type_name, method);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&dotted)),
                args,
                optional: false,
            });
        }
    }

    if let ExprKind::Ident(ref type_name) = receiver.kind {
        if type_name == "BigInteger" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args,
                optional: false,
            });
        }
    }

    if method == "toCharArray" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_to_char_array")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "formatted" {
        let mut format_args = Vec::with_capacity(args.len() + 1);
        format_args.push(Argument::positional(receiver));
        format_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_string_format")),
            args: format_args,
            optional: false,
        });
    }

    if method == "reversed" && args.is_empty() {
        return java_comparator_reversed(receiver);
    }

    if method == "thenComparing" && args.len() == 1 {
        return java_comparator_then_comparing(receiver, args[0].value.clone());
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: method,
            null_safe: false,
        })),
        args,
        optional: false,
    })
}

fn java_expr_dotted_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let mut prefix = java_expr_dotted_name(object)?;
            prefix.push('.');
            prefix.push_str(field);
            Some(prefix)
        }
        _ => None,
    }
}

fn normalise_comparator_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    match method {
        "naturalOrder" if args.is_empty() => Some(java_natural_comparator(false)),
        "reverseOrder" if args.is_empty() => Some(java_natural_comparator(true)),
        "comparing" if args.len() == 1 => Some(java_comparing_comparator(args[0].value.clone())),
        _ => None,
    }
}

fn java_lambda_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn java_two_arg_lambda(body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![java_lambda_param("__a__"), java_lambda_param("__b__")],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![],
    })
}

fn java_binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn java_ternary(cond: Expression, then_expr: Expression, else_expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(else_expr),
    })
}

fn java_call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn java_compare_expr(left: Expression, right: Expression, reverse: bool) -> Expression {
    let (less_value, greater_value) = if reverse { (1, -1) } else { (-1, 1) };
    java_ternary(
        java_binary(BinOp::Lt, left.clone(), right.clone()),
        Expression::int(less_value),
        java_ternary(
            java_binary(BinOp::Gt, left, right),
            Expression::int(greater_value),
            Expression::int(0),
        ),
    )
}

fn java_natural_comparator(reverse: bool) -> Expression {
    java_two_arg_lambda(java_compare_expr(
        Expression::ident("__a__"),
        Expression::ident("__b__"),
        reverse,
    ))
}

fn java_key_compare_expr(key_fn: Expression, left_name: &str, right_name: &str) -> Expression {
    let left_key = java_call(key_fn.clone(), vec![Expression::ident(left_name)]);
    let right_key = java_call(key_fn, vec![Expression::ident(right_name)]);
    java_compare_expr(left_key, right_key, false)
}

fn java_comparing_comparator(key_fn: Expression) -> Expression {
    java_two_arg_lambda(java_key_compare_expr(key_fn, "__a__", "__b__"))
}

fn java_comparator_call(comparator: Expression, left_name: &str, right_name: &str) -> Expression {
    java_call(
        comparator,
        vec![Expression::ident(left_name), Expression::ident(right_name)],
    )
}

fn java_comparator_reversed(comparator: Expression) -> Expression {
    java_two_arg_lambda(java_comparator_call(comparator, "__b__", "__a__"))
}

fn java_comparator_then_comparing(comparator: Expression, next: Expression) -> Expression {
    let primary_for_cond = java_comparator_call(comparator.clone(), "__a__", "__b__");
    let primary_for_result = java_comparator_call(comparator, "__a__", "__b__");
    let secondary = match &next.kind {
        ExprKind::Lambda { params, .. } if params.len() == 2 => {
            java_comparator_call(next, "__a__", "__b__")
        }
        _ => java_key_compare_expr(next, "__a__", "__b__"),
    };
    java_two_arg_lambda(java_ternary(
        java_binary(BinOp::NotEq, primary_for_cond, Expression::int(0)),
        primary_for_result,
        secondary,
    ))
}

fn java_qualified_static_type(expr: &Expression) -> Option<&str> {
    let mut parts = Vec::new();
    collect_member_chain(expr, &mut parts)?;
    if parts.len() < 2 {
        return None;
    }
    if !(parts.starts_with(&["java", "util"])
        || parts.starts_with(&["java", "lang"])
        || parts.starts_with(&["java", "time"]))
    {
        return None;
    }
    let type_name = parts.last().copied()?;
    if is_java_type_or_util(type_name) {
        Some(type_name)
    } else {
        None
    }
}

fn collect_member_chain<'a>(expr: &'a Expression, parts: &mut Vec<&'a str>) -> Option<()> {
    match expr.kind {
        ExprKind::Ident(ref name) => {
            parts.push(name.as_str());
            Some(())
        }
        ExprKind::Member {
            ref object,
            ref field,
            ..
        } => {
            collect_member_chain(object, parts)?;
            parts.push(field.as_str());
            Some(())
        }
        _ => None,
    }
}

fn is_java_type_or_util(name: &str) -> bool {
    matches!(
        name,
        "Integer"
            | "Long"
            | "Short"
            | "Byte"
            | "Float"
            | "Double"
            | "Boolean"
            | "Character"
            | "String"
            | "Math"
            | "Arrays"
            | "List"
            | "Set"
            | "Map"
            | "Collections"
            | "Objects"
            | "Optional"
            | "IntStream"
            | "LongStream"
            | "DoubleStream"
            | "Stream"
            | "Collectors"
            | "System"
            | "Thread"
            | "Runtime"
            | "Class"
            | "Comparator"
            | "Instant"
            | "Duration"
            | "ZoneId"
            | "ZoneOffset"
            | "ChronoUnit"
    )
}

fn walk_primary_atom(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("primary_atom: empty")?;
    match inner.as_rule() {
        Rule::new_expression => walk_new(inner),
        Rule::array_creation => walk_array_creation(inner),
        Rule::switch_expression => walk_switch_expression(inner),
        Rule::lambda_expression => walk_lambda(inner),
        Rule::paren_expression => walk_expression(inner.into_inner().next().ok_or("paren: empty")?),
        Rule::literal => walk_literal(inner),
        Rule::this_kw => Ok(Expression::new(ExprKind::This)),
        Rule::super_kw => Ok(Expression::new(ExprKind::Super)),
        Rule::super_method_call => walk_super_call(inner),
        Rule::class_literal => Ok(Expression::string(
            inner
                .as_str()
                .strip_suffix(".class")
                .unwrap_or(inner.as_str()),
        )),
        Rule::method_reference => walk_method_reference(inner),
        Rule::ident_name => Ok(Expression::ident(inner.as_str())),
        _ => walk_expression(inner),
    }
}

fn walk_new(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner().peekable();

    let class_name = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        extract_ref_name(&inner.next().unwrap())
    } else {
        "Object".to_string()
    };
    // skip optional type_args
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_args) {
        inner.next();
    }

    while let Some(p) = inner.next() {
        match p.as_rule() {
            Rule::argument_list => {
                let args = walk_arguments(p)?;
                if class_name.rsplit('.').next() == Some("Comparator")
                    && inner.peek().map(|next| next.as_rule()) == Some(Rule::anonymous_class_body)
                {
                    if let Some(comparator) = walk_anonymous_comparator(inner.next().unwrap())? {
                        return Ok(comparator);
                    }
                }
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args,
                }));
            }
            Rule::array_initializer => {
                // new Type[] {1, 2, 3} → array literal
                return walk_initializer_as_array(p);
            }
            Rule::array_dims => {
                let mut sizes = Vec::new();
                let mut initializer = None;
                for dim in p.into_inner() {
                    match dim.as_rule() {
                        Rule::expression => {
                            if let Ok(size) = walk_expression(dim) {
                                sizes.push(size);
                            }
                        }
                        Rule::array_initializer => initializer = Some(dim),
                        _ => {}
                    }
                }
                if sizes.is_empty() {
                    if let Some(init) = initializer {
                        return walk_initializer_as_array(init);
                    }
                }
                if sizes.len() >= 2
                    && matches!(
                        class_name.as_str(),
                        "byte"
                            | "short"
                            | "int"
                            | "long"
                            | "char"
                            | "byte[]"
                            | "short[]"
                            | "int[]"
                            | "long[]"
                            | "char[]"
                    )
                {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__new_int_2d_array")),
                        args: vec![
                            Argument::positional(sizes[0].clone()),
                            Argument::positional(sizes[1].clone()),
                        ],
                        optional: false,
                    }));
                }
                // new int[5] → __new_array(5)
                if let Some(sz) = sizes.into_iter().next() {
                    let callee = match class_name.as_str() {
                        "boolean" | "boolean[]" => "__new_bool_array",
                        "byte" | "short" | "int" | "long" | "char" | "byte[]" | "short[]"
                        | "int[]" | "long[]" | "char[]" => "__new_int_array",
                        _ => "__new_array",
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args: vec![Argument::positional(sz)],
                        optional: false,
                    }));
                }
            }
            Rule::anonymous_class_body => {
                if class_name.contains("Comparator") {
                    if let Some(comparator) = walk_anonymous_comparator(p)? {
                        return Ok(comparator);
                    }
                }
                // Anonymous class: just create new instance
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args: vec![],
                }));
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args: vec![],
    }))
}

fn walk_anonymous_comparator(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    for member in pair.into_inner() {
        if member.as_rule() != Rule::method_declaration {
            continue;
        }
        let ClassMember::Method(method) = walk_method(member)? else {
            continue;
        };
        let StmtKind::FunctionDecl {
            name, params, body, ..
        } = method.kind
        else {
            continue;
        };
        if name == "compare" && params.len() == 2 {
            return Ok(Some(Expression::new(ExprKind::Lambda {
                params,
                body: LambdaBody::Block(body),
                is_async: false,
                captures: vec![],
            })));
        }
    }
    Ok(None)
}

fn walk_array_creation(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let prim_type = inner
        .next()
        .map(|p| p.as_str().to_string())
        .unwrap_or_else(|| "Object".to_string());
    for p in inner {
        match p.as_rule() {
            Rule::array_initializer => return walk_initializer_as_array(p),
            Rule::expression => {
                let sz = walk_expression(p)?;
                let callee = match prim_type.as_str() {
                    "boolean" => "__new_bool_array",
                    "byte" | "short" | "int" | "long" | "char" => "__new_int_array",
                    _ => "__new_array",
                };
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args: vec![Argument::positional(sz)],
                    optional: false,
                }));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Array(vec![])))
}

fn walk_initializer_as_array(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut elems = Vec::new();
    for el in pair.into_inner() {
        if el.as_rule() == Rule::initializer {
            elems.push(ArrayElement {
                key: None,
                value: walk_initializer(el)?,
                spread: false,
                by_ref: false,
            });
        }
    }
    Ok(Expression::new(ExprKind::Array(elems)))
}

fn walk_super_call(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let method_name = inner
        .next()
        .ok_or("super call: missing name")?
        .as_str()
        .to_string();
    let args = if let Some(al) = inner.next() {
        if al.as_rule() == Rule::argument_list {
            walk_arguments(al)?
        } else {
            vec![]
        }
    } else {
        vec![]
    };
    Ok(Expression::new(ExprKind::SuperCall {
        method: Some(method_name),
        args,
    }))
}

fn walk_method_reference(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let obj = inner.next().ok_or("method ref: missing object")?;
    let obj_name = obj.as_str().to_string();
    let method = inner
        .next()
        .ok_or("method ref: missing method")?
        .as_str()
        .to_string();

    let obj_expr = Expression::ident(&obj_name);
    if method == "new" {
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__args__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: true,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::New {
                class: Box::new(obj_expr),
                args: vec![],
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if obj_name == "Math" {
        let callee = format!("{}.{}", obj_name, method);
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&callee)),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("Integer", "parseInt")
            | ("Integer", "valueOf")
            | ("Long", "parseLong")
            | ("Double", "parseDouble")
            | ("String", "valueOf")
    ) {
        let callee = format!("{}.{}", obj_name, method);
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&callee)),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("String", "length")
            | ("String", "toString")
            | ("String", "toUpperCase")
            | ("String", "toLowerCase")
            | ("Integer", "intValue")
            | ("Long", "longValue")
            | ("Double", "doubleValue")
            | ("Double", "intValue")
            | ("Collection", "stream")
    ) {
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("__value__")),
                    field: method,
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    Ok(Expression::new(ExprKind::Member {
        object: Box::new(obj_expr),
        field: method,
        null_safe: false,
    }))
}

fn walk_lambda(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let params_pair = inner.next().ok_or("lambda: missing params")?;
    let params = walk_lambda_params(params_pair)?;
    let body_pair = inner.next().ok_or("lambda: missing body")?;
    let body = match body_pair.as_rule() {
        Rule::function_body_block => LambdaBody::Block(walk_block(body_pair)?),
        _ => LambdaBody::Expr(Box::new(walk_expression(body_pair)?)),
    };
    Ok(Expression::new(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: vec![],
    }))
}

fn walk_lambda_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    match pair.as_rule() {
        Rule::lambda_params => {
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::typed_lambda_param_list => {
                        for tp in p.into_inner() {
                            if tp.as_rule() == Rule::typed_lambda_param {
                                let mut ti = tp.into_inner().peekable();
                                if ti.peek().map(|x| x.as_rule()) == Some(Rule::final_kw) {
                                    ti.next();
                                }
                                let type_hint =
                                    if ti.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                                        Some(extract_ref_name(&ti.next().unwrap()))
                                    } else {
                                        None
                                    };
                                let name = ti
                                    .next()
                                    .ok_or("typed lambda param: missing name")?
                                    .as_str()
                                    .to_string();
                                params.push(Param {
                                    name,
                                    type_hint,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name_list => {
                        for ip in p.into_inner() {
                            if ip.as_rule() == Rule::ident_name {
                                params.push(Param {
                                    name: ip.as_str().to_string(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name => {
                        params.push(Param {
                            name: p.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
        Rule::ident_name => {
            params.push(Param {
                name: pair.as_str().to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            });
        }
        _ => {}
    }
    Ok(params)
}

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::argument {
            let mut ai = p.into_inner();
            let first = ai.next().ok_or("arg: empty")?;
            if first.as_rule() == Rule::spread_arg {
                let e = walk_expression(first.into_inner().next().ok_or("spread: empty")?)?;
                args.push(Argument {
                    value: e,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            } else {
                args.push(Argument::positional(walk_expression(first)?));
            }
        }
    }
    Ok(args)
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("literal: empty")?;
    match inner.as_rule() {
        Rule::true_kw => Ok(Expression::bool(true)),
        Rule::false_kw => Ok(Expression::bool(false)),
        Rule::null_kw => Ok(Expression::null()),
        Rule::int_literal => {
            let s = inner.as_str().replace('_', "");
            let v = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).unwrap_or(0)
            } else if s.starts_with("0b") || s.starts_with("0B") {
                i64::from_str_radix(&s[2..], 2).unwrap_or(0)
            } else {
                s.parse::<i64>().unwrap_or(0)
            };
            Ok(Expression::int(v))
        }
        Rule::long_literal => {
            let s = inner.as_str().replace('_', "");
            let s = s.trim_end_matches(|c| c == 'l' || c == 'L');
            let v = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).unwrap_or(0)
            } else {
                s.parse::<i64>().unwrap_or(0)
            };
            Ok(Expression::int(v))
        }
        Rule::float_literal => {
            let s = inner.as_str().replace('_', "");
            let s = s.trim_end_matches(|c| matches!(c, 'f' | 'F' | 'd' | 'D'));
            Ok(Expression::float(s.parse().unwrap_or(0.0)))
        }
        Rule::char_literal => {
            let s = inner.as_str();
            let content = &s[1..s.len() - 1];
            Ok(Expression::string(&unescape_java_string(content)))
        }
        Rule::string_literal => {
            let s = inner.as_str();
            Ok(Expression::string(&unescape_java_string(
                &s[1..s.len() - 1],
            )))
        }
        Rule::text_block => {
            let s = inner.as_str();
            let content = s
                .trim_start_matches("\"\"\"")
                .trim_end_matches("\"\"\"")
                .trim_start_matches('\n');
            Ok(Expression::string(content))
        }
        _ => Ok(Expression::null()),
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Normalisations
// ════════════════════════════════════════════════════════════════════════════

/// Inject implicit `super()` into child-class constructors that don't already
/// start with an explicit super/base call.
fn inject_implicit_super(members: &mut Vec<ClassMember>) {
    for m in members.iter_mut() {
        if let ClassMember::Constructor {
            base_args, body, ..
        } = m
        {
            if base_args.is_none() {
                let already_has_super = body
                    .first()
                    .map(|s| match &s.kind {
                        StmtKind::Expr(e) => {
                            matches!(
                                &e.kind,
                                ExprKind::Call { callee, .. }
                                    if matches!(&callee.kind, ExprKind::Super)
                            ) || matches!(&e.kind, ExprKind::SuperCall { .. })
                        }
                        _ => false,
                    })
                    .unwrap_or(false);
                if !already_has_super {
                    *base_args = Some(vec![]);
                }
            }
        }
    }
}

/// Extract an explicit `super(...)` / `this(...)` call from the top of a
/// constructor body and put the args in `base_args`.
fn extract_base_call_from_body(
    body: &mut Vec<Statement>,
    base_args: &mut Option<Vec<Expression>>,
    initializer_target: &mut ConstructorInitializerTarget,
) {
    if body.is_empty() {
        return;
    }
    let target = match &body[0].kind {
        StmtKind::Expr(e) => match &e.kind {
            ExprKind::SuperCall { .. } => Some(ConstructorInitializerTarget::Base),
            ExprKind::Call { callee, .. } => match &callee.kind {
                ExprKind::Super => Some(ConstructorInitializerTarget::Base),
                ExprKind::This => Some(ConstructorInitializerTarget::This),
                _ => None,
            },
            _ => None,
        },
        _ => None,
    };
    if let Some(target) = target {
        let s = body.remove(0);
        if let StmtKind::Expr(e) = s.kind {
            let args_exprs: Vec<Expression> = match e.kind {
                ExprKind::SuperCall { args, .. } => args.into_iter().map(|a| a.value).collect(),
                ExprKind::Call { args, .. } => args.into_iter().map(|a| a.value).collect(),
                _ => vec![],
            };
            *base_args = Some(args_exprs);
            *initializer_target = target;
        }
    }
}

fn default_expr_for_java_type(type_name: &str) -> Option<Expression> {
    match type_name {
        "byte" | "short" | "int" | "long" | "char" => Some(Expression::int(0)),
        "float" | "double" => Some(Expression::float(0.0)),
        "boolean" => Some(Expression::bool(false)),
        _ => None,
    }
}

fn hoist_java_nested_types(body: &mut Vec<Statement>) {
    let mut hoisted = Vec::new();
    hoist_java_nested_types_from_stmts(body, &mut hoisted);
    body.extend(hoisted);
}

fn hoist_java_nested_types_from_stmts(body: &mut [Statement], hoisted: &mut Vec<Statement>) {
    for stmt in body {
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            let mut kept = Vec::with_capacity(members.len());
            for member in std::mem::take(members) {
                match member {
                    ClassMember::NestedType(nested) => {
                        let mut nested_body = vec![*nested];
                        hoist_java_nested_types_from_stmts(&mut nested_body, hoisted);
                        hoisted.extend(nested_body);
                    }
                    other => kept.push(other),
                }
            }
            *members = kept;
        }
    }
}

fn rewrite_java_user_tostring_calls(body: &mut [Statement]) {
    let mut tostring_classes = HashSet::new();
    collect_java_tostring_classes(body, &mut tostring_classes);
    let mut enum_values = HashMap::new();
    collect_java_enum_values(body, &mut enum_values);
    let mut double_fields = HashSet::new();
    collect_java_double_fields(body, &mut double_fields);
    rewrite_java_double_field_print_tree(body, &double_fields);
    rewrite_java_tostring_stmts(
        body,
        &tostring_classes,
        &enum_values,
        None,
        &mut HashMap::new(),
    );
}

fn collect_java_enum_values(stmts: &[Statement], out: &mut HashMap<String, Vec<String>>) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::EnumDecl { name, members, .. } => {
                out.insert(
                    name.clone(),
                    members.iter().map(|member| member.name.clone()).collect(),
                );
            }
            StmtKind::ClassDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_enum_values(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_enum_values(stmts, out),
            _ => {}
        }
    }
}

fn collect_java_double_fields(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        if let StmtKind::ClassDecl { members, .. } = &stmt.kind {
            for member in members {
                if let ClassMember::Field {
                    name,
                    type_hint: Some(type_hint),
                    ..
                } = member
                {
                    if matches!(type_hint.as_str(), "double" | "Double") {
                        out.insert(name.clone());
                    }
                }
            }
        }
    }
}

fn rewrite_java_double_field_print_tree(
    stmts: &mut [Statement],
    double_fields: &HashSet<String>,
) {
    for stmt in stmts {
        rewrite_java_double_field_prints(std::slice::from_mut(stmt), double_fields);
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            for member in members {
                match member {
                    ClassMember::Constructor { body, .. } => {
                        rewrite_java_double_field_print_tree(body, double_fields);
                    }
                    ClassMember::Method(method) => {
                        if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                            rewrite_java_double_field_print_tree(body, double_fields);
                        }
                    }
                    _ => {}
                }
            }
        }
    }
}

fn collect_java_tostring_classes(body: &[Statement], out: &mut HashSet<String>) {
    for stmt in body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            if members.iter().any(|member| {
                matches!(
                    member,
                    ClassMember::Method(method)
                        if matches!(&method.kind, StmtKind::FunctionDecl { name, .. } if name == "toString")
                )
            }) {
                out.insert(name.clone());
            }
        }
    }
}

fn rewrite_java_tostring_stmts(
    stmts: &mut [Statement],
    tostring_classes: &HashSet<String>,
    enum_values: &HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                let double_fields: std::collections::HashSet<String> = members
                    .iter()
                    .filter_map(|member| match member {
                        ClassMember::Field {
                            name,
                            type_hint: Some(type_hint),
                            ..
                        } if matches!(type_hint.as_str(), "double" | "Double") => {
                            Some(name.clone())
                        }
                        _ => None,
                    })
                    .collect();
                for member in members {
                    match member {
                        ClassMember::Constructor { params, body, .. } => {
                            let mut local_types = params
                                .iter()
                                .filter_map(|p| {
                                    p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone()))
                                })
                                .collect();
                            rewrite_java_tostring_stmts(
                                body,
                                tostring_classes,
                                enum_values,
                                Some(name),
                                &mut local_types,
                            );
                        }
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl { params, body, .. } = &mut method.kind {
                                rewrite_java_double_field_prints(body, &double_fields);
                                let mut local_types = params
                                    .iter()
                                    .filter_map(|p| {
                                        p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone()))
                                    })
                                    .collect();
                                rewrite_java_tostring_stmts(
                                    body,
                                    tostring_classes,
                                    enum_values,
                                    Some(name),
                                    &mut local_types,
                                );
                            }
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        rewrite_java_tostring_expr(
                            init,
                            tostring_classes,
                            enum_values,
                            current_class,
                            locals,
                        );
                    }
                    if let (BindingPattern::Ident(name), Some(init)) = (&decl.pattern, &decl.init) {
                        if name.starts_with("__java_switch_value_") {
                            if let ExprKind::Ident(source_name) = &init.kind {
                                if let Some(type_hint) = locals.get(source_name).cloned() {
                                    locals.insert(name.clone(), type_hint);
                                }
                            }
                        }
                    }
                    if let (BindingPattern::Ident(name), Some(type_hint)) =
                        (&decl.pattern, &decl.type_hint)
                    {
                        locals.insert(name.clone(), type_hint.clone());
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                rewrite_java_tostring_expr(value, tostring_classes, enum_values, current_class, locals);
                for target in targets {
                    rewrite_java_tostring_expr(target, tostring_classes, enum_values, current_class, locals);
                }
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                rewrite_java_tostring_expr(value, tostring_classes, enum_values, current_class, locals);
                rewrite_java_tostring_expr(target, tostring_classes, enum_values, current_class, locals);
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                rewrite_java_tostring_expr(expr, tostring_classes, enum_values, current_class, locals);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                rewrite_java_tostring_expr(cond, tostring_classes, enum_values, current_class, locals);
                rewrite_java_tostring_stmts(
                    then_body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
                for (elif_cond, elif_body) in elifs {
                    rewrite_java_tostring_expr(elif_cond, tostring_classes, enum_values, current_class, locals);
                    rewrite_java_tostring_stmts(
                        elif_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut locals.clone(),
                    );
                }
                if let Some(else_body) = else_body {
                    rewrite_java_tostring_stmts(
                        else_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut locals.clone(),
                    );
                }
            }
            StmtKind::While { cond, body, .. } => {
                rewrite_java_tostring_expr(cond, tostring_classes, enum_values, current_class, locals);
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
            }
            StmtKind::Block(body) => {
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
            }
            _ => {}
        }
    }
}

fn rewrite_java_double_field_prints(
    stmts: &mut [Statement],
    double_fields: &HashSet<String>,
) {
    for stmt in stmts {
        let StmtKind::Expr(expr) = &mut stmt.kind else {
            continue;
        };
        let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
            continue;
        };
        if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "println" || name == "print")
            || args.len() != 1
        {
            continue;
        }
        let ExprKind::Member { field, .. } = &args[0].value.kind else {
            continue;
        };
        if !double_fields.contains(field) {
            continue;
        }
        let value = args[0].value.clone();
        args[0].value = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_double_to_string")),
            args: vec![Argument::positional(value)],
            optional: false,
        });
    }
}

fn rewrite_java_tostring_expr(
    expr: &mut Expression,
    tostring_classes: &HashSet<String>,
    enum_values: &HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &HashMap<String, String>,
) {
    if let Some(replacement) = java_bigint_constant_replacement(expr) {
        *expr = replacement;
        return;
    }

    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in &mut *args {
                rewrite_java_tostring_expr(
                    &mut arg.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
            if let Some(rewritten) = rewrite_java_enum_set_static_call(callee, args, enum_values) {
                *expr = rewritten;
                return;
            }
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if field == "remove" && args.len() == 1 {
                    if let ExprKind::Call {
                        callee: key_set_callee,
                        args: key_set_args,
                        ..
                    } = &object.kind
                    {
                        if key_set_args.is_empty() {
                            if let ExprKind::Member {
                                object: map_object,
                                field: key_set_field,
                                ..
                            } = &key_set_callee.kind
                            {
                                if key_set_field == "keySet" {
                                    if let ExprKind::Ident(ref name) = map_object.kind {
                                        if java_type_is_map(locals.get(name).map(String::as_str)) {
                                            *expr = Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::ident(
                                                    "__java_map_key_set_remove",
                                                )),
                                                args: vec![
                                                    Argument::positional((**map_object).clone()),
                                                    args[0].clone(),
                                                ],
                                                optional: false,
                                            });
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                rewrite_java_tostring_expr(
                    object,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                if let ExprKind::Ident(ref name) = object.kind {
                    if java_type_is_enum_set(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_enum_set_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_uuid(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_uuid_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_instant(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_instant_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_bitset(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_bitset_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_map(locals.get(name).map(String::as_str)) {
                        if field == "keySet"
                            && matches!(
                                locals.get(name).map(String::as_str),
                                Some("TreeMap") | Some("java.util.TreeMap")
                            )
                        {
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__java_sorted_map_key_set")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false,
                            });
                            return;
                        }
                        if let Some(internal) = java_map_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                }
                if matches!(field.as_str(), "add" | "offer" | "poll")
                    && matches!(
                        object.kind,
                        ExprKind::Ident(ref name)
                            if matches!(
                                locals.get(name).map(String::as_str),
                                Some("TreeSet") | Some("PriorityQueue")
                            )
                    )
                {
                    let internal = if field == "poll" {
                        "__java_sorted_poll"
                    } else {
                        "__java_sorted_add"
                    };
                    let mut new_args = Vec::with_capacity(args.len() + 1);
                    new_args.push(Argument::positional((**object).clone()));
                    new_args.extend(args.iter().cloned());
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: new_args,
                        optional: false,
                    });
                    return;
                }
                if let Some(name) = java_bigint_method_name(field) {
                    if java_expr_is_bigint(object, locals) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(name)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if field == "toString"
                    && java_expr_has_user_tostring(object, tostring_classes, current_class, locals)
                {
                    *field = "tostring".to_string();
                }
            } else {
                rewrite_java_tostring_expr(
                    callee,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
        }
        ExprKind::Member { object, field, .. } => {
            if matches!(field.as_str(), "SECONDS" | "MILLIS")
                && java_member_chain_ends_with(object, "ChronoUnit")
            {
                *expr = Expression::string(field);
                return;
            }
            rewrite_java_tostring_expr(object, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_tostring_expr(object, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(index, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_tostring_expr(left, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(right, tostring_classes, enum_values, current_class, locals);
            rewrite_java_switch_enum_label(left, right, locals);
            rewrite_java_switch_enum_label(right, left, locals);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_tostring_expr(inner, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_tostring_expr(value, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(target, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_tostring_expr(cond, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(then, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(else_, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_tostring_expr(
                    &mut elem.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_java_tostring_expr(
                    &mut arg.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
        }
        _ => {}
    }
}

fn rewrite_java_switch_enum_label(
    maybe_switch_value: &Expression,
    maybe_label: &mut Expression,
    locals: &std::collections::HashMap<String, String>,
) {
    let switch_type = match &maybe_switch_value.kind {
        ExprKind::Ident(name) if name.starts_with("__java_switch_value_") => locals.get(name),
        _ => None,
    };
    let Some(type_hint) = switch_type else {
        return;
    };
    let ExprKind::Ident(label) = &maybe_label.kind else {
        return;
    };
    if label.starts_with("__") {
        return;
    }
    *maybe_label = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(java_type_simple_name(type_hint))),
        field: label.clone(),
        null_safe: false,
    });
}

fn rewrite_java_enum_set_static_call(
    callee: &Expression,
    args: &[Argument],
    enum_values: &HashMap<String, Vec<String>>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !java_member_chain_ends_with(object, "EnumSet") {
        return None;
    }
    let names = match field.as_str() {
        "noneOf" | "allOf" => {
            let enum_name = args.first().and_then(|arg| java_string_literal(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "of" => {
            let enum_name = args
                .first()
                .and_then(|arg| java_enum_type_from_member_expr(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "range" => {
            let enum_name = args
                .first()
                .and_then(|arg| java_enum_type_from_member_expr(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "copyOf" | "complementOf" => Expression::null(),
        _ => return None,
    };
    let internal = match field.as_str() {
        "noneOf" => "__java_enum_set_none_of",
        "allOf" => "__java_enum_set_all_of",
        "of" => "__java_enum_set_of",
        "copyOf" => "__java_enum_set_copy_of",
        "complementOf" => "__java_enum_set_complement_of",
        "range" => "__java_enum_set_range",
        _ => return None,
    };
    let mut new_args = Vec::with_capacity(args.len() + 1);
    match field.as_str() {
        "copyOf" | "complementOf" => new_args.extend(args.iter().cloned()),
        "of" | "range" => {
            new_args.push(Argument::positional(names));
            new_args.extend(args.iter().cloned());
        }
        _ => new_args.push(Argument::positional(names)),
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(internal)),
        args: new_args,
        optional: false,
    }))
}

fn java_member_chain_ends_with(expr: &Expression, expected: &str) -> bool {
    let mut parts = Vec::new();
    collect_member_chain(expr, &mut parts).is_some() && parts.last().copied() == Some(expected)
}

fn java_string_literal(expr: &Expression) -> Option<&str> {
    if let ExprKind::Lit(Literal::Str(value)) = &expr.kind {
        Some(value.as_str())
    } else {
        None
    }
}

fn java_enum_type_from_member_expr(expr: &Expression) -> Option<&str> {
    if let ExprKind::Member { object, .. } = &expr.kind {
        if let ExprKind::Ident(name) = &object.kind {
            return Some(name.as_str());
        }
    }
    None
}

fn java_enum_names_expr(
    enum_values: &HashMap<String, Vec<String>>,
    enum_name: &str,
) -> Option<Expression> {
    let base = enum_name.rsplit('.').next().unwrap_or(enum_name);
    let values = enum_values.get(base)?;
    Some(Expression::new(ExprKind::Array(
        values
            .iter()
            .map(|name| ArrayElement {
                key: None,
                value: Expression::string(name),
                spread: false,
                by_ref: false,
            })
            .collect(),
    )))
}

fn java_type_is_enum_set(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "EnumSet"
}

fn java_enum_set_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "add" => "__java_enum_set_add",
        "addAll" => "__java_enum_set_add_all",
        "contains" => "__java_enum_set_contains",
        "containsAll" => "__java_enum_set_contains_all",
        "remove" => "__java_enum_set_remove",
        "equals" => "__java_enum_set_equals",
        "hashCode" => "__java_enum_set_hash_code",
        "iterator" => "__java_enum_set_iterator",
        "getClass" => "__java_enum_set_get_class",
        _ => return None,
    })
}

fn java_type_is_map(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    matches!(
        base,
        "Map"
            | "HashMap"
            | "LinkedHashMap"
            | "WeakHashMap"
            | "TreeMap"
            | "SortedMap"
            | "NavigableMap"
            | "Hashtable"
    )
}

fn java_map_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "get" => "__java_map_get",
        "put" => "__java_map_put",
        "putAll" => "__java_map_put_all",
        "remove" => "__java_map_remove",
        "getOrDefault" => "__java_map_get_or_default",
        "containsKey" => "__java_map_contains_key",
        "containsValue" => "__java_map_contains_value",
        "keySet" => "__java_map_key_set",
        "values" => "__java_map_values",
        "entrySet" => "__java_map_entry_set",
        "putIfAbsent" => "__java_map_put_if_absent",
        "computeIfAbsent" => "__java_map_compute_if_absent",
        "computeIfPresent" => "__java_map_compute_if_present",
        "compute" => "__java_map_compute",
        "merge" => "__java_map_merge",
        "replace" => "__java_map_replace",
        "replaceAll" => "__java_map_replace_all",
        "forEach" => "__java_map_for_each",
        "clear" => "__java_map_clear",
        "size" => "__java_map_size",
        "isEmpty" => "__java_map_is_empty",
        "equals" => "__java_map_equals",
        "firstEntry" => "__java_sorted_map_first_entry",
        "lastEntry" => "__java_sorted_map_last_entry",
        "ceilingEntry" => "__java_sorted_map_ceiling_entry",
        "floorEntry" => "__java_sorted_map_floor_entry",
        "higherEntry" => "__java_sorted_map_higher_entry",
        "lowerEntry" => "__java_sorted_map_lower_entry",
        "subMap" => "__java_map_sub_map",
        "headMap" => "__java_map_head_map",
        "tailMap" => "__java_map_tail_map",
        _ => return None,
    })
}

fn java_type_is_bitset(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "BitSet"
}

fn java_type_is_uuid(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "UUID"
}

fn java_type_is_instant(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "Instant"
}

fn java_instant_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_instant_compare_to",
        "equals" => "__java_instant_equals",
        "toString" => "__java_instant_to_string",
        "hashCode" => "__java_instant_hash_code",
        _ => return None,
    })
}

fn java_uuid_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_uuid_compare_to",
        "hashCode" => "__java_uuid_hash_code",
        _ => return None,
    })
}

fn java_bitset_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "set" => "__java_bitset_set",
        "get" => "__java_bitset_get",
        "clear" => "__java_bitset_clear",
        "flip" => "__java_bitset_flip",
        "cardinality" => "__java_bitset_cardinality",
        "length" => "__java_bitset_length",
        "size" => "__java_bitset_size",
        "isEmpty" => "__java_bitset_is_empty",
        "nextSetBit" => "__java_bitset_next_set_bit",
        "nextClearBit" => "__java_bitset_next_clear_bit",
        "previousSetBit" => "__java_bitset_previous_set_bit",
        "previousClearBit" => "__java_bitset_previous_clear_bit",
        "and" => "__java_bitset_and",
        "or" => "__java_bitset_or",
        "xor" => "__java_bitset_xor",
        "andNot" => "__java_bitset_and_not",
        "intersects" => "__java_bitset_intersects",
        "equals" => "__java_bitset_equals",
        "clone" => "__java_bitset_clone",
        "stream" => "__java_bitset_stream",
        "toLongArray" | "toByteArray" => "__java_bitset_to_array",
        "toString" => "__java_bitset_to_string",
        "hashCode" => "__java_bitset_hash_code",
        _ => return None,
    })
}

fn java_bigint_method_name(method: &str) -> Option<&'static str> {
    match method {
        "toString" => Some("__java_bigint_to_string"),
        "add" => Some("__java_bigint_add"),
        "subtract" => Some("__java_bigint_subtract"),
        "multiply" => Some("__java_bigint_multiply"),
        "mod" => Some("__java_bigint_mod"),
        "gcd" => Some("__java_bigint_gcd"),
        "pow" => Some("__java_bigint_pow"),
        "compareTo" => Some("__java_bigint_compare_to"),
        "negate" => Some("__java_bigint_negate"),
        "abs" => Some("__java_bigint_abs"),
        "signum" => Some("__java_bigint_signum"),
        "max" => Some("__java_bigint_max"),
        "min" => Some("__java_bigint_min"),
        "bitLength" => Some("__java_bigint_bit_length"),
        "testBit" => Some("__java_bigint_test_bit"),
        "shiftLeft" => Some("__java_bigint_shift_left"),
        "shiftRight" => Some("__java_bigint_shift_right"),
        "and" => Some("__java_bigint_and"),
        "or" => Some("__java_bigint_or"),
        "xor" => Some("__java_bigint_xor"),
        "not" => Some("__java_bigint_not"),
        "isProbablePrime" => Some("__java_bigint_is_probable_prime"),
        "nextProbablePrime" => Some("__java_bigint_next_probable_prime"),
        _ => None,
    }
}

fn java_expr_is_bigint(
    expr: &Expression,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(name)
            .is_some_and(|type_hint| java_type_simple_name(type_hint) == "BigInteger"),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => java_type_simple_name(name) == "BigInteger",
            ExprKind::Member { .. } => {
                java_qualified_static_type(class).is_some_and(|name| name == "BigInteger")
            }
            _ => false,
        },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                name.starts_with("__java_bigint") || name == "BigInteger.valueOf"
            }
            _ => false,
        },
        _ => java_bigint_constant_replacement(expr).is_some(),
    }
}

fn java_bigint_constant_replacement(expr: &Expression) -> Option<Expression> {
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        let is_bigint_type = java_qualified_static_type(object)
            .is_some_and(|name| name == "BigInteger")
            || java_expr_dotted_name(object).as_deref() == Some("java.math.BigInteger");
        if is_bigint_type {
            let value = match field.as_str() {
                "ZERO" => "0",
                "ONE" => "1",
                "TEN" => "10",
                _ => return None,
            };
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args: vec![Argument::positional(Expression::int(
                    value.parse::<i64>().unwrap_or(0),
                ))],
                optional: false,
            }));
        }
    }
    None
}

fn java_type_simple_name(type_name: &str) -> &str {
    type_name.rsplit('.').next().unwrap_or(type_name)
}

fn java_expr_has_user_tostring(
    expr: &Expression,
    tostring_classes: &std::collections::HashSet<String>,
    current_class: Option<&str>,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(name)
            .is_some_and(|type_hint| tostring_classes.contains(type_hint)),
        ExprKind::This => current_class.is_some_and(|name| tostring_classes.contains(name)),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => tostring_classes.contains(name),
            _ => false,
        },
        _ => false,
    }
}

fn normalize_java_class_tree(body: &mut [Statement]) {
    use std::collections::HashMap;

    let mut class_members = HashMap::new();
    collect_java_class_member_names(body, &mut class_members);
    normalize_java_class_tree_with_members(body, &class_members);
}

fn collect_java_class_member_names(
    body: &[Statement],
    out: &mut std::collections::HashMap<String, JavaClassMemberNames>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                out.insert(name.clone(), JavaClassMemberNames::from_members(members));
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_class_member_names(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_class_member_names(stmts, out),
            _ => {}
        }
    }
}

fn normalize_java_class_tree_with_members(
    body: &mut [Statement],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                members,
                ..
            } => {
                let mut names = class_members.get(name).cloned().unwrap_or_default();
                for parent in parents {
                    if let Some(parent_names) = class_members.get(parent) {
                        names.fields.extend(parent_names.fields.iter().cloned());
                        names.methods.extend(parent_names.methods.iter().cloned());
                    }
                }
                normalize_java_class_members(members, &names);
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        normalize_java_class_tree_with_members(
                            std::slice::from_mut(nested),
                            class_members,
                        );
                    }
                }
            }
            StmtKind::Block(stmts) => normalize_java_class_tree_with_members(stmts, class_members),
            _ => {}
        }
    }
}

#[derive(Clone, Default)]
struct JavaClassMemberNames {
    fields: std::collections::HashSet<String>,
    methods: std::collections::HashSet<String>,
}

impl JavaClassMemberNames {
    fn from_members(members: &[ClassMember]) -> Self {
        let fields = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Field {
                    name, modifiers, ..
                } if !modifiers.is_static => Some(name.clone()),
                _ => None,
            })
            .collect();
        let methods = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Method(func) => match &func.kind {
                    StmtKind::FunctionDecl {
                        name, modifiers, ..
                    } if !modifiers.is_static => Some(name.clone()),
                    _ => None,
                },
                _ => None,
            })
            .collect();
        Self { fields, methods }
    }
}

fn normalize_java_class_members(members: &mut [ClassMember], names: &JavaClassMemberNames) {
    if names.fields.is_empty() && names.methods.is_empty() {
        return;
    }

    for member in members {
        match member {
            ClassMember::Constructor { params, body, .. } => {
                let mut locals = params.iter().map(|p| p.name.clone()).collect();
                normalize_java_stmts(body, &names.fields, &names.methods, &mut locals);
            }
            ClassMember::Method(func) => {
                if let StmtKind::FunctionDecl {
                    params,
                    body,
                    modifiers,
                    ..
                } = &mut func.kind
                {
                    if modifiers.is_static {
                        continue;
                    }
                    let mut locals = params.iter().map(|p| p.name.clone()).collect();
                    normalize_java_stmts(body, &names.fields, &names.methods, &mut locals);
                }
            }
            _ => {}
        }
    }
}

fn normalize_java_stmts(
    stmts: &mut [Statement],
    fields: &std::collections::HashSet<String>,
    methods: &std::collections::HashSet<String>,
    locals: &mut std::collections::HashSet<String>,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        normalize_java_expr(init, fields, methods, locals, false);
                    }
                    collect_binding_names(&decl.pattern, locals);
                }
            }
            StmtKind::Assign { targets, value } => {
                normalize_java_expr(value, fields, methods, locals, false);
                for target in targets {
                    normalize_java_expr(target, fields, methods, locals, true);
                }
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                normalize_java_expr(value, fields, methods, locals, false);
                normalize_java_expr(target, fields, methods, locals, true);
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                normalize_java_expr(expr, fields, methods, locals, false);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                normalize_java_expr(cond, fields, methods, locals, false);
                normalize_java_stmts(then_body, fields, methods, &mut locals.clone());
                for (elif_cond, elif_body) in elifs {
                    normalize_java_expr(elif_cond, fields, methods, locals, false);
                    normalize_java_stmts(elif_body, fields, methods, &mut locals.clone());
                }
                if let Some(else_body) = else_body {
                    normalize_java_stmts(else_body, fields, methods, &mut locals.clone());
                }
            }
            StmtKind::While { cond, body, .. } => {
                normalize_java_expr(cond, fields, methods, locals, false);
                normalize_java_stmts(body, fields, methods, &mut locals.clone());
            }
            StmtKind::Block(body) => {
                normalize_java_stmts(body, fields, methods, &mut locals.clone());
            }
            _ => {}
        }
    }
}

fn collect_binding_names(pattern: &BindingPattern, locals: &mut std::collections::HashSet<String>) {
    match pattern {
        BindingPattern::Ident(name) => {
            locals.insert(name.clone());
        }
        BindingPattern::Object(props) => {
            for prop in props {
                if let Some(value) = &prop.value {
                    collect_binding_names(value, locals);
                } else {
                    locals.insert(prop.key.clone());
                }
            }
        }
        BindingPattern::Array(elems) => {
            for elem in elems {
                match elem {
                    ArrayPatternElem::Pattern(pattern, _) => collect_binding_names(pattern, locals),
                    ArrayPatternElem::Rest(name) => {
                        locals.insert(name.clone());
                    }
                    ArrayPatternElem::Hole => {}
                }
            }
        }
    }
}

fn normalize_java_expr(
    expr: &mut Expression,
    fields: &std::collections::HashSet<String>,
    methods: &std::collections::HashSet<String>,
    locals: &std::collections::HashSet<String>,
    is_assignment_target: bool,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if fields.contains(name) && !locals.contains(name) => {
            let field = name.clone();
            expr.kind = ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field,
                null_safe: false,
            };
        }
        ExprKind::Call { callee, args, .. } => {
            for arg in args {
                normalize_java_expr(&mut arg.value, fields, methods, locals, false);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if methods.contains(name) && !locals.contains(name) {
                    let method = name.clone();
                    callee.kind = ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: method,
                        null_safe: false,
                    };
                    return;
                }
            }
            normalize_java_expr(callee, fields, methods, locals, false);
        }
        ExprKind::Member { object, .. } => {
            normalize_java_expr(object, fields, methods, locals, false);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_java_expr(object, fields, methods, locals, false);
            normalize_java_expr(index, fields, methods, locals, false);
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_java_expr(left, fields, methods, locals, false);
            normalize_java_expr(right, fields, methods, locals, false);
        }
        ExprKind::Unary { expr: inner, .. } => {
            normalize_java_expr(inner, fields, methods, locals, is_assignment_target);
        }
        ExprKind::Assign { target, value } => {
            normalize_java_expr(value, fields, methods, locals, false);
            normalize_java_expr(target, fields, methods, locals, true);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_java_expr(cond, fields, methods, locals, false);
            normalize_java_expr(then, fields, methods, locals, false);
            normalize_java_expr(else_, fields, methods, locals, false);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                normalize_java_expr(&mut elem.value, fields, methods, locals, false);
            }
        }
        _ => {}
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

/// Extract a simple type name from a `type_ref` or `ref_type` node.
fn extract_ref_name(pair: &Pair<Rule>) -> String {
    match pair.as_rule() {
        Rule::type_ref => {
            let dims = pair
                .clone()
                .into_inner()
                .filter(|p| p.as_rule() == Rule::dim_suffix)
                .count();
            for p in pair.clone().into_inner() {
                match p.as_rule() {
                    Rule::primitive_type => return format!("{}{}", p.as_str(), "[]".repeat(dims)),
                    Rule::ref_type => {
                        return format!("{}{}", extract_ref_name(&p), "[]".repeat(dims));
                    }
                    _ => {}
                }
            }
            let base = pair
                .as_str()
                .split('<')
                .next()
                .unwrap_or("Object")
                .trim()
                .trim_end_matches("[]")
                .to_string();
            format!("{}{}", base, "[]".repeat(dims))
        }
        Rule::ref_type => {
            for p in pair.clone().into_inner() {
                if p.as_rule() == Rule::qualified_name {
                    return p
                        .as_str()
                        .rsplit('.')
                        .next()
                        .unwrap_or(p.as_str())
                        .to_string();
                }
            }
            pair.as_str()
                .split('<')
                .next()
                .unwrap_or("Object")
                .trim()
                .to_string()
        }
        _ => pair
            .as_str()
            .split('<')
            .next()
            .unwrap_or("Object")
            .trim()
            .to_string(),
    }
}

fn str_to_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        ">>>" => BinOp::UShr,
        _ => BinOp::Add,
    }
}

fn compound_op_to_binop(s: &str) -> BinOp {
    match s.trim_end_matches('=') {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        ">>>" => BinOp::UShr,
        _ => BinOp::Add,
    }
}

fn unescape_java_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('u') => {
                    let mut hex = String::with_capacity(4);
                    while matches!(chars.clone().next(), Some('u')) {
                        chars.next();
                    }
                    for _ in 0..4 {
                        if let Some(h) = chars.next() {
                            hex.push(h);
                        }
                    }
                    if let Ok(code) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(code) {
                            out.push(ch);
                        }
                    }
                }
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('"') => out.push('"'),
                Some('\'') => out.push('\''),
                Some('\\') => out.push('\\'),
                Some('0') => out.push('\0'),
                Some(c) => {
                    out.push('\\');
                    out.push(c);
                }
                None => {}
            }
        } else {
            out.push(c);
        }
    }
    out
}

fn find_main_class(body: &[Statement]) -> Option<String> {
    for stmt in body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            for m in members {
                if let ClassMember::Method(func) = m {
                    if let StmtKind::FunctionDecl {
                        name: fname,
                        modifiers,
                        ..
                    } = &func.kind
                    {
                        if fname == "main" && modifiers.is_static {
                            return Some(name.clone());
                        }
                    }
                }
            }
        }
    }
    None
}
