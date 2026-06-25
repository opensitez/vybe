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
use pest::Parser;
use pest::iterators::Pair;

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
        Rule::final_kw | Rule::static_kw | Rule::public_kw | Rule::private_kw
            | Rule::protected_kw | Rule::abstract_kw | Rule::synchronized_kw
            | Rule::native_kw | Rule::transient_kw | Rule::volatile_kw
            | Rule::strictfp_kw | Rule::default_kw | Rule::sealed_kw
            | Rule::non_sealed_kw | Rule::var_kw
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
                members.push(ClassMember::NestedType(Box::new(Statement::new(walk_class(p)?))));
            }
            Rule::interface_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(walk_interface(p)?))));
            }
            Rule::enum_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(walk_enum_decl(p)?))));
            }
            Rule::record_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(walk_record(p)?))));
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

    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::throws_clause => {}
            Rule::function_body_block => {
                body = walk_block(p)?;
                // Extract super(...) or this(...) call from top of body
                extract_base_call_from_body(&mut body, &mut base_args);
            }
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        params,
        body,
        base_args,
        initializer_target: ConstructorInitializerTarget::Base,
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
                None
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
            let e = pair.into_inner().find(|p| !is_kw(p.as_rule())).map(walk_expression).transpose()?;
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
            let e = pair.into_inner().find(|p| !is_kw(p.as_rule())).map(walk_expression).transpose()?;
            StmtKind::Return(e)
        }

        Rule::labeled_statement => {
            // Strip label; walk inner statement
            let inner = pair.into_inner().find(|p| !matches!(p.as_rule(), Rule::ident_name));
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
            StmtKind::Expr(Expression::new(ExprKind::SuperCall {
                method: None,
                args,
            }))
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
            let cls = pair.into_inner().find(|p| p.as_rule() == Rule::class_declaration);
            if let Some(c) = cls {
                walk_class(c)?
            } else {
                return Ok(None);
            }
        }

        Rule::local_record_declaration => {
            let rec = pair.into_inner().find(|p| p.as_rule() == Rule::record_declaration);
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

    let kind = if is_final { VarDeclKind::Const } else { VarDeclKind::Let };

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

fn walk_var_declarator(pair: Pair<Rule>, type_hint: Option<String>) -> Result<VarDeclarator, String> {
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
    let mut inner = pair.into_inner().peekable();

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
        let kind = if is_final { VarDeclKind::Const } else { VarDeclKind::Let };
        let mut decls = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::var_declarator {
                decls.push(walk_var_declarator(p, type_hint.clone())?);
            }
        }
        return Ok(Statement::new(StmtKind::VarDecl { declarations: decls, kind }));
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
        Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Sequence(exprs)))))
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

    let var = inner.next().ok_or("for-each: missing var")?.as_str().to_string();
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
    let mut inner = pair.into_inner();
    let switch_expr = walk_expr_inner(&mut inner)?;

    let mut cases: Vec<SwitchCase> = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for case_pair in inner {
        if case_pair.as_rule() != Rule::switch_case {
            continue;
        }
        let mut ci = case_pair.into_inner().peekable();
        let mut conditions: Vec<CaseCondition> = Vec::new();
        let mut body: Vec<Statement> = Vec::new();
        let mut is_default = false;
        let src = {
            let tmp = ci.peek().map(|p| p.as_str().to_string()).unwrap_or_default();
            tmp
        };

        if src.trim() == "default" {
            is_default = true;
            ci.next(); // consume "default"
        }

        for p in ci {
            match p.as_rule() {
                Rule::switch_label => {
                    if let Some(expr_p) = p.into_inner().next() {
                        if let Ok(e) = walk_expression(expr_p) {
                            conditions.push(CaseCondition::Value(e));
                        }
                    }
                }
                Rule::switch_rule_body => {
                    for rb in p.into_inner() {
                        if let Some(s) = walk_statement(rb)? {
                            body.push(s);
                        }
                    }
                }
                _ => {
                    if let Some(s) = walk_statement(p)? {
                        body.push(s);
                    }
                }
            }
        }

        if is_default || conditions.is_empty() {
            default = Some(body);
        } else {
            cases.push(SwitchCase { conditions, body });
        }
    }

    Ok(StmtKind::Switch {
        expr: switch_expr,
        cases,
        default,
    })
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
    let name = inner.next().ok_or("param: missing name")?.as_str().to_string();
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
            // (Type)expr — erase cast, return expr
            let mut ci = pair.into_inner();
            let _cast_type = ci.next(); // cast_type
            if let Some(operand) = ci.next() {
                walk_expression(operand)
            } else {
                Ok(Expression::null())
            }
        }
        Rule::postfix_expression => walk_postfix(pair),
        Rule::primary_chain => walk_primary_chain(pair),
        Rule::primary_atom => walk_primary_atom(pair),
        Rule::lambda_expression => walk_lambda(pair),
        Rule::switch_expression => Ok(Expression::null()), // simplification
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
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(rhs),
        });
    }
    Ok(left)
}

fn walk_instanceof(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut base = walk_expression(inner.next().ok_or("instanceof: empty")?)?;

    let mut saw_instanceof = false;
    for p in inner {
        match p.as_rule() {
            Rule::instanceof_kw => { saw_instanceof = true; }
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
                let method_name = ci.next().ok_or("method call: missing name")?.as_str().to_string();
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
    if let ExprKind::Member { object: ref root_obj, field: ref root_field, .. } = receiver.kind {
        if let ExprKind::Ident(ref root_name) = root_obj.kind {
            if root_name == "System" && root_field == "out"
                && (method == "println" || method == "print" || method == "printf")
            {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(&method)),
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

    // Static type method calls: Integer.parseInt("42") → call "Integer.parseInt"
    // The profile has dotted builtins like "Integer.parseInt", "Math.max", etc.
    if let ExprKind::Ident(ref type_name) = receiver.kind {
        if is_java_type_or_util(type_name) {
            let dotted = format!("{}.{}", type_name, method);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&dotted)),
                args,
                optional: false,
            });
        }
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

fn is_java_type_or_util(name: &str) -> bool {
    matches!(
        name,
        "Integer" | "Long" | "Short" | "Byte" | "Float" | "Double"
        | "Boolean" | "Character" | "String" | "Math" | "Arrays"
        | "Collections" | "Objects" | "Optional" | "Stream"
        | "System" | "Thread" | "Runtime" | "Class"
    )
}

fn walk_primary_atom(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("primary_atom: empty")?;
    match inner.as_rule() {
        Rule::new_expression => walk_new(inner),
        Rule::array_creation => walk_array_creation(inner),
        Rule::switch_expression => Ok(Expression::null()),
        Rule::lambda_expression => walk_lambda(inner),
        Rule::paren_expression => {
            walk_expression(inner.into_inner().next().ok_or("paren: empty")?)
        }
        Rule::literal => walk_literal(inner),
        Rule::this_kw => Ok(Expression::new(ExprKind::This)),
        Rule::super_kw => Ok(Expression::new(ExprKind::Super)),
        Rule::super_method_call => walk_super_call(inner),
        Rule::class_literal => Ok(Expression::null()),
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

    for p in inner {
        match p.as_rule() {
            Rule::argument_list => {
                let args = walk_arguments(p)?;
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
                // new int[5] → __new_array(5)
                if let Some(size_p) = p.into_inner().next() {
                    if let Ok(sz) = walk_expression(size_p) {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__new_array")),
                            args: vec![Argument::positional(sz)],
                            optional: false,
                        }));
                    }
                }
            }
            Rule::anonymous_class_body => {
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

fn walk_array_creation(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let _prim_type = inner.next();
    for p in inner {
        match p.as_rule() {
            Rule::array_initializer => return walk_initializer_as_array(p),
            Rule::expression => {
                let sz = walk_expression(p)?;
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__new_array")),
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
    let method_name = inner.next().ok_or("super call: missing name")?.as_str().to_string();
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
    let method = inner.next().ok_or("method ref: missing method")?.as_str().to_string();

    let obj_expr = walk_expression(obj)?;
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
                                if ti.peek().map(|x| x.as_rule()) == Some(Rule::final_kw) { ti.next(); }
                                let type_hint = if ti.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                                    Some(extract_ref_name(&ti.next().unwrap()))
                                } else { None };
                                let name = ti.next().ok_or("typed lambda param: missing name")?.as_str().to_string();
                                params.push(Param {
                                    name, type_hint, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name_list => {
                        for ip in p.into_inner() {
                            if ip.as_rule() == Rule::ident_name {
                                params.push(Param {
                                    name: ip.as_str().to_string(), type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name => {
                        params.push(Param {
                            name: p.as_str().to_string(), type_hint: None, default: None,
                            pass_by: PassBy::Value, is_rest: false,
                            is_kwargs: false, is_optional: false, is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
        Rule::ident_name => {
            params.push(Param {
                name: pair.as_str().to_string(), type_hint: None, default: None,
                pass_by: PassBy::Value, is_rest: false,
                is_kwargs: false, is_optional: false, is_nullable: false,
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
                args.push(Argument { value: e, name: None, by_ref: false, spread: true });
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
            let ch = if content.starts_with('\\') {
                match content.chars().nth(1) {
                    Some('n') => '\n',
                    Some('t') => '\t',
                    Some('r') => '\r',
                    Some('\'') => '\'',
                    Some('\\') => '\\',
                    Some('0') => '\0',
                    _ => content.chars().nth(1).unwrap_or('\0'),
                }
            } else {
                content.chars().next().unwrap_or('\0')
            };
            Ok(Expression::int(ch as i64))
        }
        Rule::string_literal => {
            let s = inner.as_str();
            Ok(Expression::string(&unescape_java_string(&s[1..s.len() - 1])))
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
        if let ClassMember::Constructor { base_args, body, .. } = m {
            if base_args.is_none() {
                let already_has_super = body
                    .first()
                    .map(|s| match &s.kind {
                        StmtKind::Expr(e) => matches!(
                            &e.kind,
                            ExprKind::Call { callee, .. }
                                if matches!(&callee.kind, ExprKind::Super)
                        ) || matches!(&e.kind, ExprKind::SuperCall { .. }),
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
fn extract_base_call_from_body(body: &mut Vec<Statement>, base_args: &mut Option<Vec<Expression>>) {
    if body.is_empty() {
        return;
    }
    let is_super = match &body[0].kind {
        StmtKind::Expr(e) => {
            matches!(&e.kind, ExprKind::SuperCall { .. })
                || matches!(&e.kind, ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Super))
        }
        _ => false,
    };
    if is_super {
        let s = body.remove(0);
        if let StmtKind::Expr(e) = s.kind {
            let args_exprs: Vec<Expression> = match e.kind {
                ExprKind::SuperCall { args, .. } => {
                    args.into_iter().map(|a| a.value).collect()
                }
                ExprKind::Call { args, .. } => {
                    args.into_iter().map(|a| a.value).collect()
                }
                _ => vec![],
            };
            *base_args = Some(args_exprs);
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

/// Extract a simple type name from a `type_ref` or `ref_type` node.
fn extract_ref_name(pair: &Pair<Rule>) -> String {
    match pair.as_rule() {
        Rule::type_ref => {
            for p in pair.clone().into_inner() {
                match p.as_rule() {
                    Rule::primitive_type => return p.as_str().to_string(),
                    Rule::ref_type => return extract_ref_name(&p),
                    _ => {}
                }
            }
            pair.as_str()
                .split('<')
                .next()
                .unwrap_or("Object")
                .trim()
                .to_string()
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
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('"') => out.push('"'),
                Some('\'') => out.push('\''),
                Some('\\') => out.push('\\'),
                Some('0') => out.push('\0'),
                Some(c) => { out.push('\\'); out.push(c); }
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
