//! Go walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//!
//! ## Go-specific normalisations
//!
//! - **Multiple return values**: Go functions can return multiple values.
//!   For simplicity we compile to returning a single array/tuple.
//! - **Short variable declaration** (`:=`): Maps to `VarDecl` with `Let`.
//! - **Methods**: Go methods on structs are compiled as regular functions
//!   with the receiver as the first parameter.
//! - **Structs**: Mapped to `ClassDecl` with fields.
//! - **Interfaces**: Mapped to `InterfaceDecl`.
//! - **`range`**: Mapped to `ForIn` with `of: true`.
//! - **`defer`**: Currently ignored (no-op) — Go's defer semantics require
//!   runtime support not yet available.
//! - **`go`**: Currently ignored (no-op) — goroutines require runtime support.
//! - **`fallthrough`**: Not yet supported in switch.
//! - **`select`**: Not yet supported.
//! - **`chan` / `<-`**: Not yet supported.
//! - **`nil`**: Mapped to `ExprKind::Lit(Literal::Null)`.
//! - **`make` / `new`**: `make` for slices/maps is rewritten to array/dict
//!   creation. `new(T)` becomes `&T{}` (pointer to zero value).
//! - **`append`**: Rewritten to array push.
//! - **`len` / `cap`**: Builtin functions mapped to host calls.
//! - **`panic` / `recover`**: Mapped to throw/try-catch.
//! - **`_` blank identifier**: Ignored in assignments.

use pest::Parser;
use pest::iterators::Pair;
use crate::ast::*;
use super::{GoParser, Rule};

// ══════════════════════════════════════════════════════════════════════════════════════════
// Entry point
// ══════════════════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = GoParser::parse(Rule::program, source)
        .map_err(|e| format!("Go parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut _package_name = String::new();

    for top in pairs {
        if top.as_rule() == Rule::EOI { continue; }
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            _ => {
                if let Some(stmt) = walk_top_level(top)? {
                    body.push(stmt);
                }
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI => continue,
                Rule::package_clause => {
                    _package_name = walk_package_clause(pair)?;
                }
                Rule::import_declarations => {
                    for imp in pair.into_inner() {
                        if imp.as_rule() == Rule::import_declaration {
                            imports.push(walk_import(imp)?);
                        }
                    }
                }
                _ => {
                    if let Some(stmt) = walk_top_level(pair)? {
                        body.push(stmt);
                    }
                }
            }
        }
    }

    Ok(Module {
        name: _package_name,
        language: Lang::Go,
        body,
        imports,
    })
}

fn walk_package_clause(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let mut path = String::new();
    let mut alias: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::import_spec => {
                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => {
                            alias = Some(spec_inner.as_str().to_string());
                        }
                        Rule::string_literal => {
                            path = unquote(spec_inner.as_str());
                        }
                        _ => {}
                    }
                }
            }
            Rule::string_literal => {
                path = unquote(inner.as_str());
            }
            _ => {}
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span: Span::default(),
    })
}

fn unquote(s: &str) -> String {
    if s.len() >= 2 && ((s.starts_with('"') && s.ends_with('"')) || (s.starts_with('\'') && s.ends_with('\'')) || (s.starts_with('`') && s.ends_with('`'))) {
        s[1..s.len()-1].to_string()
    } else {
        s.to_string()
    }
}

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::function_declaration => Ok(Some(walk_function_decl(pair)?)),
        Rule::method_declaration => Ok(Some(walk_method_decl(pair)?)),
        Rule::var_declaration => Ok(Some(walk_var_decl(pair)?)),
        Rule::const_declaration => Ok(Some(walk_const_decl(pair)?)),
        Rule::type_declaration => walk_type_decl(pair),
        Rule::declaration => {
            for inner in pair.into_inner() {
                return walk_top_level(inner);
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

// ── Function declarations ─────────────────────────────────────────────────────────────

fn walk_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::signature => {
                let (p, rt) = walk_signature(inner)?;
                params = p;
                return_type = rt;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_method_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut receiver_name = String::new();
    let mut receiver_type = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::receiver => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::ident_name => receiver_name = r_inner.as_str().to_string(),
                        Rule::type_annotation => receiver_type = walk_type(r_inner),
                        _ => {}
                    }
                }
            }
            Rule::ident_name => method_name = inner.as_str().to_string(),
            Rule::signature => {
                let (p, rt) = walk_signature(inner)?;
                params = p;
                return_type = rt;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    // Prepend receiver as first parameter
    params.insert(0, Param {
        name: if receiver_name.is_empty() { "self".to_string() } else { receiver_name },
        type_hint: Some(receiver_type),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    });

    Ok(Statement::new(StmtKind::FunctionDecl {
        name: method_name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_signature(pair: Pair<Rule>) -> Result<(Vec<Param>, Option<String>), String> {
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::parameter_list => {
                params = walk_parameter_list(inner)?;
            }
            Rule::result => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::type_annotation => return_type = Some(walk_type(r_inner)),
                        Rule::parameter_list => {
                            // Multiple return values — represent as array
                            let p = walk_parameter_list(r_inner)?;
                            return_type = Some(format!("[{}]", p.len()));
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_annotation => {
                return_type = Some(walk_type(inner));
            }
            _ => {}
        }
    }

    Ok((params, return_type))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter_decl {
            let mut names = Vec::new();
            let mut type_hint: Option<String> = None;

            for p_inner in inner.into_inner() {
                match p_inner.as_rule() {
                    Rule::ident_name => names.push(p_inner.as_str().to_string()),
                    Rule::ident_list => {
                        for id in p_inner.into_inner() {
                            if id.as_rule() == Rule::ident_name {
                                names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::type_annotation => type_hint = Some(walk_type(p_inner)),
                    _ => {}
                }
            }

            for name in names {
                params.push(Param {
                    name,
                    type_hint: type_hint.clone(),
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
    Ok(params)
}

fn walk_type(pair: Pair<Rule>) -> String {
    pair.as_str().to_string()
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block_statement | Rule::function_body => {
                stmts.append(&mut walk_block(inner)?);
            }
            Rule::statement_list => {
                for s in inner.into_inner() {
                    if s.as_rule() == Rule::statement {
                        stmts.push(walk_statement(s)?);
                    }
                }
            }
            Rule::statement => {
                stmts.push(walk_statement(inner)?);
            }
            _ => {}
        }
    }
    Ok(stmts)
}

// ── Variable declarations ─────────────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::var_spec | Rule::const_spec => {
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Let)?;
                declarations.append(&mut decls);
            }
            Rule::var_group | Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::var_spec || spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Let)?;
                        declarations.append(&mut decls);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    }))
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::const_spec => {
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Const)?;
                declarations.append(&mut decls);
            }
            Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Const)?;
                        declarations.append(&mut decls);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const,
    }))
}

fn walk_var_spec(pair: Pair<Rule>, _kind: VarDeclKind) -> Result<(Vec<VarDeclarator>, Option<String>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init: Option<Expression> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if id.as_rule() == Rule::ident_name {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => {
                let exprs = walk_expression_list(inner)?;
                if !exprs.is_empty() {
                    init = Some(exprs.into_iter().next().unwrap());
                }
            }
            Rule::expression => {
                init = Some(walk_expression(inner)?);
            }
            _ => {}
        }
    }

    let mut declarations = Vec::new();
    for name in names {
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: init.clone(),
            type_hint: type_hint.clone(),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, type_hint))
}

// ── Type declarations (struct, interface, type alias) ─────────────────────────────────

fn walk_type_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_spec => {
                let mut name = String::new();
                let mut type_str = String::new();

                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => name = spec_inner.as_str().to_string(),
                        Rule::type_annotation => type_str = walk_type(spec_inner),
                        Rule::struct_type => {
                            return Ok(Some(walk_struct_type(name, spec_inner)?));
                        }
                        Rule::interface_type => {
                            return Ok(Some(walk_interface_type(name, spec_inner)?));
                        }
                        _ => {}
                    }
                }

                // Type alias — just create a variable with the type name
                if !type_str.is_empty() && !name.is_empty() {
                    return Ok(Some(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(name),
                            init: Some(Expression::new(ExprKind::Lit(Literal::Str(type_str)))),
                            type_hint: None,
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    })));
                }
            }
            Rule::type_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::type_spec {
                        return walk_type_decl(spec.into());
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_struct_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::field_decl {
            let mut field_names = Vec::new();
            let mut field_type: Option<String> = None;

            for f_inner in inner.into_inner() {
                match f_inner.as_rule() {
                    Rule::ident_list => {
                        for id in f_inner.into_inner() {
                            if id.as_rule() == Rule::ident_name {
                                field_names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::ident_name => field_names.push(f_inner.as_str().to_string()),
                    Rule::type_annotation => field_type = Some(walk_type(f_inner)),
                    _ => {}
                }
            }

            for fname in field_names {
                    members.push(ClassMember::Field {
                    name: fname,
                    type_hint: field_type.clone(),
                    init: None,
                    modifiers: Modifiers::default(),
                    with_events: false,
                    array_bounds: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: Vec::new(),
    }))
}

fn walk_interface_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::interface_member {
            let mut method_name = String::new();
            let mut params = Vec::new();
            let mut return_type: Option<String> = None;

            for m_inner in inner.into_inner() {
                match m_inner.as_rule() {
                    Rule::ident_name => method_name = m_inner.as_str().to_string(),
                    Rule::signature => {
                        let (p, rt) = walk_signature(m_inner)?;
                        params = p;
                        return_type = rt;
                    }
                    _ => {}
                }
            }

            if !method_name.is_empty() {
                members.push(InterfaceMember::Method {
                    name: method_name,
                    params,
                    return_type,
                    is_sub: false,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::InterfaceDecl {
        name,
        parents: Vec::new(),
        members,
        decorators: Vec::new(),
    }))
}

// ── Statements ─────────────────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let rule = pair.as_rule();
    if rule == Rule::statement {
        if let Some(inner) = pair.into_inner().next() {
            return walk_statement(inner);
        }
        return Ok(Statement::new(StmtKind::Empty));
    }

    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => StmtKind::Block(walk_block(pair)?),
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::assignment_statement => walk_assignment(pair)?,
        Rule::short_var_declaration => walk_short_var_decl(pair)?,
        Rule::inc_dec_statement => walk_inc_dec(pair)?,
        Rule::var_declaration => walk_var_decl(pair)?.kind,
        Rule::const_declaration => walk_const_decl(pair)?.kind,
        Rule::if_statement => walk_if(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::goto_statement => StmtKind::GoTo(walk_goto(pair)?),
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::defer_statement => StmtKind::Empty, // TODO: implement defer
        Rule::go_statement => StmtKind::Empty,    // TODO: implement goroutines
        Rule::send_statement => StmtKind::Empty,  // TODO: implement channels
        _ => StmtKind::Empty,
    };
    Ok(Statement::new(kind))
}

fn walk_assignment(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut op = "=";
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression_list => {
                if targets.is_empty() {
                    targets = walk_expression_list(inner)?;
                } else {
                    values = walk_expression_list(inner)?;
                }
            }
            Rule::assign_op => op = inner.as_str(),
            _ => {}
        }
    }

    if op != "=" {
        // Compound assignment
        if targets.len() == 1 && values.len() == 1 {
            let compound_op = match op {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "%=" => CompoundOp::Mod,
                _ => CompoundOp::Add,
            };
            return Ok(StmtKind::CompoundAssign {
                target: targets.into_iter().next().unwrap(),
                op: compound_op,
                value: values.into_iter().next().unwrap(),
            });
        }
    }

    if values.len() == 1 {
        Ok(StmtKind::Assign {
            targets,
            value: values.into_iter().next().unwrap(),
        })
    } else if !values.is_empty() {
        // Multiple values — pack into array
        let arr_elems: Vec<ArrayElement> = values.into_iter().map(|v| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Array(arr_elems)),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_short_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if id.as_rule() == Rule::ident_name {
                        let name = id.as_str().to_string();
                        if name != "_" {
                            names.push(name);
                        }
                    }
                }
            }
            Rule::expression_list => {
                values = walk_expression_list(inner)?;
            }
            _ => {}
        }
    }

    let mut declarations = Vec::new();
    let value = if values.len() == 1 {
        values.into_iter().next().unwrap()
    } else if !values.is_empty() {
        let arr_elems: Vec<ArrayElement> = values.into_iter().map(|v| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Expression::new(ExprKind::Array(arr_elems))
    } else {
        Expression::new(ExprKind::Lit(Literal::Null))
    };

    for name in names {
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: Some(value.clone()),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        });
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_inc_dec(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let mut is_inc = true;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {
                let s = inner.as_str();
                if s == "--" { is_inc = false; }
            }
        }
    }

    if let Some(target) = expr {
        Ok(StmtKind::CompoundAssign {
            target,
            op: if is_inc { CompoundOp::Add } else { CompoundOp::Sub },
            value: Expression::new(ExprKind::Lit(Literal::Int(1))),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(inner)?);
                }
            }
            Rule::block_statement => {
                if then_body.is_empty() {
                    then_body = walk_block(inner)?;
                }
            }
            Rule::else_clause => {
                for e_inner in inner.into_inner() {
                    match e_inner.as_rule() {
                        Rule::block_statement => else_body = Some(walk_block(e_inner)?),
                        Rule::if_statement => {
                            let elif = walk_if(e_inner)?;
                            if let StmtKind::If { cond: c, then_body: t, else_body: e, .. } = elif {
                                then_body.push(Statement::new(StmtKind::If {
                                    cond: c,
                                    then_body: t,
                                    elifs: Vec::new(),
                                    else_body: e,
                                }));
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            _ => {}
        }
    }

    let mut then = then_body;
    if let Some(pre) = pre_stmt {
        then.insert(0, *pre);
    }

    Ok(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body: then,
        elifs: Vec::new(),
        else_body,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::expr_case_clause => {
                let mut conditions: Vec<CaseCondition> = Vec::new();
                let mut body = Vec::new();

                for c_inner in inner.into_inner() {
                    match c_inner.as_rule() {
                        Rule::expr_switch_case => {
                            for sc_inner in c_inner.into_inner() {
                                if sc_inner.as_rule() == Rule::expression_list {
                                    for expr in walk_expression_list(sc_inner)? {
                                        conditions.push(CaseCondition::Value(expr));
                                    }
                                } else if sc_inner.as_rule() == Rule::kw_default {
                                    // default case
                                }
                            }
                        }
                        Rule::statement_list => {
                            body = walk_statement_list(c_inner)?;
                        }
                        _ => {}
                    }
                }

                if conditions.is_empty() {
                    default = Some(body);
                } else {
                    cases.push(SwitchCase { conditions, body });
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        cases,
        default,
    })
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement {
            stmts.push(walk_statement(inner)?);
        }
    }
    Ok(stmts)
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body = Vec::new();
    let mut is_range = false;
    let mut range_vars = Vec::new();
    let mut range_iter = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::for_clause => {
                for fc_inner in inner.into_inner() {
                    match fc_inner.as_rule() {
                        Rule::short_var_declaration => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::expression_statement => {
                            let expr = walk_expression(first_meaningful(fc_inner)?)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            } else if update.is_none() {
                                update = Some(expr);
                            }
                        }
                        Rule::assignment_statement => {
                            let assign = walk_assignment(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(assign)));
                            } else if update.is_none() {
                                if let StmtKind::Assign { targets, value } = assign {
                                    if let Some(target) = targets.into_iter().next() {
                                        update = Some(Expression::new(ExprKind::Assign {
                                            target: Box::new(target),
                                            value: Box::new(value),
                                        }));
                                    }
                                }
                            }
                        }
                        Rule::inc_dec_statement => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(inc_dec)));
                            } else if update.is_none() {
                                if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                    let bin_op = match op {
                                        CompoundOp::Add => BinOp::Add,
                                        CompoundOp::Sub => BinOp::Sub,
                                        CompoundOp::Mul => BinOp::Mul,
                                        CompoundOp::Div => BinOp::Div,
                                        CompoundOp::Mod => BinOp::Mod,
                                        _ => BinOp::Add,
                                    };
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target.clone()),
                                        value: Box::new(Expression::new(ExprKind::Binary {
                                            op: bin_op,
                                            left: Box::new(target),
                                            right: Box::new(value),
                                        })),
                                    }));
                                }
                            }
                        }
                        Rule::for_inc_dec => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                let bin_op = match op {
                                    CompoundOp::Add => BinOp::Add,
                                    CompoundOp::Sub => BinOp::Sub,
                                    CompoundOp::Mul => BinOp::Mul,
                                    CompoundOp::Div => BinOp::Div,
                                    CompoundOp::Mod => BinOp::Mod,
                                    _ => BinOp::Add,
                                };
                                update = Some(Expression::new(ExprKind::Assign {
                                    target: Box::new(target.clone()),
                                    value: Box::new(Expression::new(ExprKind::Binary {
                                        op: bin_op,
                                        left: Box::new(target),
                                        right: Box::new(value),
                                    })),
                                }));
                            }
                        }
                        Rule::for_assign_nosemi => {
                            let assign = walk_assignment(fc_inner)?;
                            if let StmtKind::Assign { targets, value } = assign {
                                if let Some(target) = targets.into_iter().next() {
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target),
                                        value: Box::new(value),
                                    }));
                                }
                            }
                        }
                        Rule::expression => {
                            if cond.is_none() {
                                cond = Some(walk_expression(fc_inner)?);
                            } else if update.is_none() {
                                update = Some(walk_expression(fc_inner)?);
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::range_clause => {
                is_range = true;
                for rc_inner in inner.into_inner() {
                    match rc_inner.as_rule() {
                        Rule::expression_list => {
                            for expr in walk_expression_list(rc_inner)? {
                                let name = if let ExprKind::Ident(id) = &expr.kind {
                                    id.clone()
                                } else {
                                    "_".to_string()
                                };
                                range_vars.push(BindingPattern::Ident(name));
                            }
                        }
                        Rule::ident_list => {
                            for id in rc_inner.into_inner() {
                                if id.as_rule() == Rule::ident_name {
                                    range_vars.push(BindingPattern::Ident(id.as_str().to_string()));
                                }
                            }
                        }
                        Rule::expression => {
                            range_iter = Some(walk_expression(rc_inner)?);
                        }
                        _ => {}
                    }
                }
            }
            Rule::expression => {
                cond = Some(walk_expression(inner)?);
            }
            Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    if is_range {
        let var = range_vars.get(0).cloned().unwrap_or_else(|| BindingPattern::Ident("_".to_string()));
        let var_name = match var {
            BindingPattern::Ident(name) => name,
            _ => "_".to_string(),
        };
        let key = if range_vars.len() > 1 {
            let key_pat = range_vars.get(1).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else {
            None
        };

        Ok(StmtKind::ForIn {
            var: var_name,
            key,
            iter: range_iter.unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new()))),
            body,
            of: true,
            else_body: None,
            is_async: false,
        })
    } else if init.is_none() && update.is_none() && cond.is_some() {
        Ok(StmtKind::While {
            cond: cond.unwrap(),
            body,
            else_body: None,
        })
    } else {
        Ok(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression_list {
            values = walk_expression_list(inner)?;
        } else if inner.as_rule() == Rule::expression {
            values.push(walk_expression(inner)?);
        }
    }

    if values.len() == 1 {
        Ok(StmtKind::Return(Some(values.into_iter().next().unwrap())))
    } else if values.len() > 1 {
        let arr_elems: Vec<ArrayElement> = values.into_iter().map(|v| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(StmtKind::Return(Some(Expression::new(ExprKind::Array(arr_elems)))))
    } else {
        Ok(StmtKind::Return(None))
    }
}

fn walk_goto(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut label = String::new();
    let mut stmt = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => label = inner.as_str().to_string(),
            Rule::statement => stmt = Some(walk_statement(inner)?),
            _ => {}
        }
    }

    if let Some(s) = stmt {
        Ok(StmtKind::Block(vec![
            Statement::new(StmtKind::Label(label)),
            s,
        ]))
    } else {
        Ok(StmtKind::Label(label))
    }
}

// ── Expressions ─────────────────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::expression {
        let mut left = None;
        let mut ops = Vec::new();

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::unary_expression => {
                    if left.is_none() {
                        left = Some(walk_unary_expression(inner)?);
                    } else {
                        ops.push((None, walk_unary_expression(inner)?));
                    }
                }
                Rule::binary_op => {
                    ops.push((Some(inner.as_str().to_string()), Expression::new(ExprKind::Lit(Literal::Null))));
                }
                _ => {}
            }
        }

        if let Some(mut result) = left {
            let mut i = 0;
            while i < ops.len() {
                if let (Some(op), _) = &ops[i] {
                    if i + 1 < ops.len() {
                        let right = ops[i + 1].1.clone();
                        let bin_op = parse_bin_op(op);
                        result = Expression::new(ExprKind::Binary {
                            op: bin_op,
                            left: Box::new(result),
                            right: Box::new(right),
                        });
                        i += 2;
                        continue;
                    }
                }
                i += 1;
            }
            return Ok(result);
        }
    } else if pair.as_rule() == Rule::unary_expression {
        return walk_unary_expression(pair);
    } else if pair.as_rule() == Rule::primary {
        return walk_primary(pair);
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut op = None;
    let mut operand = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::unary_op => op = Some(inner.as_str().to_string()),
            Rule::unary_expression => operand = Some(walk_unary_expression(inner)?),
            Rule::primary => operand = Some(walk_primary(inner)?),
            _ => {}
        }
    }

    if let Some(uop) = op {
        let un_op = match uop.as_str() {
            "-" => UnaryOp::Neg,
            "!" => UnaryOp::Not,
            "+" => UnaryOp::Pos,
            "^" => UnaryOp::BitNot,
            "*" => UnaryOp::Deref,
            "&" => UnaryOp::AddrOf,
            "<-" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))), // channel receive — not supported
            _ => UnaryOp::Pos,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: un_op,
            expr: Box::new(operand.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        }))
    } else {
        operand.ok_or_else(|| "Empty unary expression".to_string())
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut base = None;
    let mut chain = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::operand => {
                base = Some(walk_operand(inner)?);
            }
            Rule::selector => {
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::ident_name {
                        chain.push(PrimaryChain::Member(s_inner.as_str().to_string()));
                    }
                }
            }
            Rule::index => {
                for i_inner in inner.into_inner() {
                    if i_inner.as_rule() == Rule::expression {
                        chain.push(PrimaryChain::Index(walk_expression(i_inner)?));
                    }
                }
            }
            Rule::two_index_slice | Rule::three_index_slice => {
                let mut start = None;
                let mut end = None;
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::expression {
                        if start.is_none() {
                            start = Some(walk_expression(s_inner)?);
                        } else if end.is_none() {
                            end = Some(walk_expression(s_inner)?);
                        }
                    }
                }
                chain.push(PrimaryChain::Slice { start, end });
            }
            Rule::call => {
                let mut args = Vec::new();
                for c_inner in inner.into_inner() {
                    if c_inner.as_rule() == Rule::argument_list {
                        for arg_inner in c_inner.into_inner() {
                            if arg_inner.as_rule() == Rule::argument {
                                let mut spread = false;
                                let mut val = None;
                                for expr_inner in arg_inner.into_inner() {
                                    if expr_inner.as_rule() == Rule::expression {
                                        val = Some(walk_expression(expr_inner)?);
                                    } else if expr_inner.as_str() == "..." {
                                        spread = true;
                                    }
                                }
                                if let Some(expr) = val {
                                    args.push(Argument {
                                        value: expr,
                                        name: None,
                                        by_ref: false,
                                        spread,
                                    });
                                }
                            }
                        }
                    }
                }
                chain.push(PrimaryChain::Call(args));
            }
            Rule::type_assertion => {
                // type assertions like .(Type) — ignore for now
            }
            _ => {}
        }
    }

    if let Some(mut result) = base {
        for item in chain {
            result = match item {
                PrimaryChain::Member(name) => Expression::new(ExprKind::Member {
                    object: Box::new(result),
                    field: name,
                    null_safe: false,
                }),
                PrimaryChain::Index(idx) => Expression::new(ExprKind::Index {
                    object: Box::new(result),
                    index: Box::new(idx),
                    null_safe: false,
                }),
                PrimaryChain::Slice { start, end } => {
                    let start_expr = start.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
                    let end_expr = end.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(result),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument { value: start_expr, name: None, by_ref: false, spread: false },
                            Argument { value: end_expr, name: None, by_ref: false, spread: false },
                        ],
                        optional: false,
                    })
                }
                PrimaryChain::Call(args) => Expression::new(ExprKind::Call {
                    callee: Box::new(result),
                    args,
                    optional: false,
                }),
            };
        }
        Ok(result)
    } else {
        Ok(Expression::new(ExprKind::Lit(Literal::Null)))
    }
}

#[derive(Clone)]
enum PrimaryChain {
    Member(String),
    Index(Expression),
    Slice { start: Option<Expression>, end: Option<Expression> },
    Call(Vec<Argument>),
}

fn walk_operand(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal => return walk_literal(inner),
            Rule::ident_name => {
                let name = inner.as_str();
                // Go builtins
                match name {
                    "nil" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))),
                    "true" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
                    "false" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                    _ => return Ok(Expression::new(ExprKind::Ident(name.to_string()))),
                }
            }
            Rule::expression => return walk_expression(inner),
            Rule::composite_literal => return walk_composite_literal(inner),
            Rule::function_literal => return walk_function_literal(inner),
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::numeric_literal => {
                let s = inner.as_str().replace('_', "");
                if s.starts_with("0x") || s.starts_with("0X") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 16) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 2) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 8) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.contains('.') || s.contains('e') || s.contains('E') || s.contains('p') || s.contains('P') {
                    if let Ok(f) = s.parse::<f64>() {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Float(f))));
                    }
                } else if let Ok(n) = s.parse::<i64>() {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                }
            }
            Rule::string_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(inner.as_str())))));
            }
            Rule::bool_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Bool(inner.as_str() == "true"))));
            }
            Rule::nil_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_composite_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::new();
    let mut elements = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal_type => {
                for lt_inner in inner.into_inner() {
                    match lt_inner.as_rule() {
                        Rule::ident_name => type_name = lt_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            let t = walk_type(lt_inner);
                            if t.starts_with("map[") {
                                type_name = "map".to_string();
                            } else if t.starts_with("[]") || t.starts_with("[") {
                                type_name = "[]".to_string();
                            } else {
                                type_name = t;
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::literal_value => {
                for lv_inner in inner.into_inner() {
                    if lv_inner.as_rule() == Rule::element_list {
                        elements = walk_element_list(lv_inner)?;
                    }
                }
            }
            _ => {}
        }
    }

    if type_name == "map" || type_name.starts_with("map[") {
        // Build a dict/object literal
        let mut props = Vec::new();
        for (key, val) in elements {
            let key_str = match &key.kind {
                ExprKind::Lit(Literal::Str(s)) => s.clone(),
                ExprKind::Ident(s) => s.clone(),
                ExprKind::Lit(Literal::Int(n)) => n.to_string(),
                _ => format!("{:?}", key),
            };
            props.push(ObjectProperty::KeyValue {
                key: Expression::new(ExprKind::Lit(Literal::Str(key_str))),
                value: val,
            });
        }
        Ok(Expression::new(ExprKind::Object(props)))
    } else {
        // Array literal or struct literal
        let arr_elems: Vec<ArrayElement> = elements.into_iter().map(|(_, v)| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(Expression::new(ExprKind::Array(arr_elems)))
    }
}

fn walk_element_list(pair: Pair<Rule>) -> Result<Vec<(Expression, Expression)>, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::keyed_element {
            let mut key = None;
            let mut val = None;

            for ke_inner in inner.into_inner() {
                match ke_inner.as_rule() {
                    Rule::ident_name => {
                        if key.is_none() {
                            key = Some(Expression::new(ExprKind::Ident(ke_inner.as_str().to_string())));
                        } else if val.is_none() {
                            val = Some(Expression::new(ExprKind::Ident(ke_inner.as_str().to_string())));
                        }
                    }
                    Rule::expression => {
                        if key.is_none() {
                            key = Some(walk_expression(ke_inner)?);
                        } else if val.is_none() {
                            val = Some(walk_expression(ke_inner)?);
                        }
                    }
                    Rule::element => {
                        for e_inner in ke_inner.into_inner() {
                            if e_inner.as_rule() == Rule::expression {
                                if val.is_none() {
                                    val = Some(walk_expression(e_inner)?);
                                }
                            } else if e_inner.as_rule() == Rule::literal_value {
                                val = Some(Expression::new(ExprKind::Lit(Literal::Null)));
                            }
                        }
                    }
                    Rule::literal_value => {
                        val = Some(Expression::new(ExprKind::Lit(Literal::Null)));
                    }
                    _ => {}
                }
            }

            let value = val.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            if let Some(k) = key {
                elements.push((k, value));
            } else {
                elements.push((Expression::new(ExprKind::Lit(Literal::Null)), value));
            }
        }
    }
    Ok(elements)
}

fn walk_function_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::signature => {
                let (p, _) = walk_signature(inner)?;
                params = p;
            }
            Rule::function_body | Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    }))
}

fn walk_expression_list(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression {
            exprs.push(walk_expression(inner)?);
        }
    }
    Ok(exprs)
}

// ── Helpers ───────────────────────────────────────────────────────────────────────────────

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() != Rule::EOI {
            return Ok(inner);
        }
    }
    Err("No meaningful child".to_string())
}

fn parse_bin_op(op: &str) -> BinOp {
    match op {
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
        "&^" => BinOp::BitAnd, // Go's bit clear — map to BitAnd as approximation
        _ => BinOp::Add,
    }
}
