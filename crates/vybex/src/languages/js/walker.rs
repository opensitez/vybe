use pest::Parser;
use pest::iterators::Pair;
use super::{JsParser, Rule};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = JsParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    // pest wraps everything in the `program` rule — unwrap it
    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => { body.push(walk_statement(top)?); continue; }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                Rule::import_statement => imports.push(walk_import(pair)?),
                _ => body.push(walk_statement(pair)?),
            }
        }
    }
    // JS function hoisting: function declarations are visible before their
    // textual position. Reorder so they come first in the body. This mirrors
    // what the JS engine does at parse time — function decls are hoisted to
    // the top of their enclosing scope.
    let mut hoisted = Vec::new();
    let mut rest = Vec::new();
    for stmt in body {
        if matches!(stmt.kind, StmtKind::FunctionDecl { .. }) {
            hoisted.push(stmt);
        } else {
            rest.push(stmt);
        }
    }
    hoisted.append(&mut rest);
    let body = hoisted;

    Ok(Module {
        name: "main".into(),
        language: Lang::JavaScript,
        body,
        imports,
    })
}

// ── Statements ──────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => {
            let stmts = pair.into_inner()
                .filter(|p| p.as_rule() != Rule::NEWLINE)
                .map(walk_statement)
                .collect::<Result<Vec<_>, _>>()?;
            StmtKind::Block(stmts)
        }
        Rule::variable_declaration => walk_var_decl(pair)?,
        Rule::function_declaration | Rule::async_function_declaration => walk_func_decl(pair)?,
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::if_statement => walk_if(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::while_statement => walk_while(pair)?,
        Rule::do_while_statement => walk_do_while(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => walk_break(pair)?,
        Rule::continue_statement => walk_continue(pair)?,
        Rule::throw_statement => walk_throw(pair)?,
        Rule::try_statement => walk_try(pair)?,
        Rule::export_statement => walk_export(pair)?,
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::NEWLINE => return Ok(Statement::new(StmtKind::Empty)),
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

// ── Variable declaration ────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let kind_pair = next_rule(&mut inner, Rule::var_kind)?;
    let var_kind = match kind_pair.as_str() {
        "var" => VarDeclKind::Var,
        "let" => VarDeclKind::Let,
        "const" => VarDeclKind::Const,
        _ => VarDeclKind::Let,
    };
    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            declarations.push(walk_var_declarator(p)?);
        }
    }
    Ok(StmtKind::VarDecl { declarations, kind: var_kind })
}

fn walk_var_declarator(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut inner = pair.into_inner();
    let pattern = walk_binding_pattern(inner.next().ok_or("Expected binding pattern")?)?;
    let init = inner.next().map(walk_expression).transpose()?;
    Ok(VarDeclarator {
        pattern,
        type_hint: None,
        init,
        array_bounds: None,
        with_events: false,
    })
}

fn walk_binding_pattern(pair: Pair<Rule>) -> Result<BindingPattern, String> {
    match pair.as_rule() {
        Rule::ident_name => Ok(BindingPattern::Ident(pair.as_str().to_string())),
        Rule::binding_pattern => walk_binding_pattern(pair.into_inner().next().ok_or("Empty binding")?),
        Rule::object_pattern => {
            let props = pair.into_inner()
                .filter(|p| p.as_rule() == Rule::object_pattern_prop)
                .map(walk_object_pattern_prop)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BindingPattern::Object(props))
        }
        Rule::array_pattern => {
            let elems = pair.into_inner()
                .filter(|p| p.as_rule() == Rule::array_pattern_elem)
                .map(walk_array_pattern_elem)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BindingPattern::Array(elems))
        }
        other => Err(format!("Unexpected binding pattern: {:?}", other)),
    }
}

fn walk_object_pattern_prop(pair: Pair<Rule>) -> Result<ObjectPatternProp, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty object pattern prop")?;
    let key = first.as_str().to_string();
    let mut value = None;
    let mut default = None;
    for p in inner {
        match p.as_rule() {
            Rule::binding_pattern => value = Some(walk_binding_pattern(p)?),
            _ => default = Some(walk_expression(p)?),
        }
    }
    Ok(ObjectPatternProp { key, value, default })
}

fn walk_array_pattern_elem(pair: Pair<Rule>) -> Result<ArrayPatternElem, String> {
    let src = pair.as_str().to_string();
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty array pattern elem")?;
    match first.as_rule() {
        Rule::array_hole => Ok(ArrayPatternElem::Hole),
        Rule::ident_name => {
            // Could be rest (...name) or simple binding
            let name = first.as_str().to_string();
            let default = inner.next().map(walk_expression).transpose()?;
            // If parent started with "..." it's rest — check source text
            if src.starts_with("...") {
                Ok(ArrayPatternElem::Rest(name))
            } else {
                Ok(ArrayPatternElem::Pattern(BindingPattern::Ident(name), default))
            }
        }
        Rule::binding_pattern => {
            let pat = walk_binding_pattern(first)?;
            let default = inner.next().map(walk_expression).transpose()?;
            Ok(ArrayPatternElem::Pattern(pat, default))
        }
        other => Err(format!("Unexpected array pattern elem: {:?}", other)),
    }
}

// ── Function declaration ────────────────────────────────────────────────────

/// Recursively scan a function body for `yield` / `yield from` expressions.
/// Does NOT descend into nested function/closure/class bodies — those are
/// their own generator scope.
fn body_contains_yield(stmts: &[Statement]) -> bool {
    fn ey(e: &Expression) -> bool {
        match &e.kind {
            ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
            // Scope boundaries — separate generator context
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => false,
            // Leaves
            ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::This | ExprKind::Super
            | ExprKind::AddressOf(_) | ExprKind::Destructure(_) => false,
            // Unary wrappers
            ExprKind::Unary { expr: i, .. } | ExprKind::IsType { expr: i, .. }
            | ExprKind::Cast { expr: i, .. } | ExprKind::TypeOf(i)
            | ExprKind::Spread(i) | ExprKind::Await(i) | ExprKind::Void(i)
            | ExprKind::Delete(i) => ey(i),
            // Binary / two-child
            ExprKind::Binary { left: a, right: b, .. }
            | ExprKind::NullCoalesce { left: a, right: b }
            | ExprKind::Assign { target: a, value: b }
            | ExprKind::Walrus { target: a, value: b }
            | ExprKind::Range { start: a, end: b, .. } => ey(a) || ey(b),
            ExprKind::StaticAccess { class: a, member: b } => ey(a) || ey(b),
            ExprKind::Ternary { cond, then, else_ } => ey(cond) || ey(then) || ey(else_),
            ExprKind::Member { object, .. } => ey(object),
            ExprKind::Index { object, index } => ey(object) || ey(index),
            ExprKind::Call { callee, args, .. } => ey(callee) || args.iter().any(|a| ey(&a.value)),
            ExprKind::New { class, args } => ey(class) || args.iter().any(|a| ey(&a.value)),
            ExprKind::SuperCall { args, .. } => args.iter().any(|a| ey(&a.value)),
            ExprKind::Array(elems) => elems.iter().any(|el| ey(&el.value) || el.key.as_ref().map_or(false, |k| ey(k))),
            ExprKind::Tuple(es) | ExprKind::Set(es) | ExprKind::Sequence(es) => es.iter().any(|x| ey(x)),
            ExprKind::Object(props) => props.iter().any(|p| match p {
                ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => ey(key) || ey(value),
                ObjectProperty::Spread(x) => ey(x),
                _ => false,
            }),
            ExprKind::Interpolation(parts) => parts.iter().any(|p| match p {
                InterpolPart::Expr(x) | InterpolPart::Formatted(x, _) => ey(x),
                _ => false,
            }),
            ExprKind::Match { subject, arms } => {
                ey(subject) || arms.iter().any(|a| {
                    a.conditions.as_ref().map_or(false, |cs| cs.iter().any(|c| ey(c)))
                    || ey(&a.body)
                })
            }
            ExprKind::Comprehension { element, generators, .. } => {
                ey(element) || generators.iter().any(|g| ey(&g.iter) || g.conditions.iter().any(|c| ey(c)))
            }
            ExprKind::Slice { lower, upper, step } => {
                [lower, upper, step].iter().any(|o| o.as_ref().map_or(false, |x| ey(x)))
            }
        }
    }
    fn sy(s: &Statement) -> bool {
        match &s.kind {
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => false,
            StmtKind::Expr(e) => ey(e),
            StmtKind::Block(ss) => ss.iter().any(|s| sy(s)),
            StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|d| d.init.as_ref().map_or(false, |e| ey(e))),
            StmtKind::Return(e) => e.as_ref().map_or(false, |e| ey(e)),
            StmtKind::If { cond, then_body, elifs, else_body } => {
                ey(cond) || then_body.iter().any(|s| sy(s))
                || elifs.iter().any(|(c, b)| ey(c) || b.iter().any(|s| sy(s)))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::While { cond, body, else_body } => {
                ey(cond) || body.iter().any(|s| sy(s))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::DoWhile { body, cond, .. } => body.iter().any(|s| sy(s)) || ey(cond),
            StmtKind::For { init, cond, update, body } => {
                init.as_ref().map_or(false, |s| sy(s))
                || cond.as_ref().map_or(false, |e| ey(e))
                || update.as_ref().map_or(false, |e| ey(e))
                || body.iter().any(|s| sy(s))
            }
            StmtKind::ForIn { iter, body, else_body, .. } => {
                ey(iter) || body.iter().any(|s| sy(s))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Switch { expr, cases, default } => {
                ey(expr) || cases.iter().any(|c| c.body.iter().any(|s| sy(s)))
                || default.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                body.iter().any(|s| sy(s))
                || catches.iter().any(|c| c.body.iter().any(|s| sy(s)))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
                || finally.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Assign { targets, value } => targets.iter().any(|e| ey(e)) || ey(value),
            StmtKind::CompoundAssign { target, value, .. } => ey(target) || ey(value),
            StmtKind::Throw { expr, cause } => {
                expr.as_ref().map_or(false, |e| ey(e)) || cause.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::Labeled { body, .. } => sy(body),
            StmtKind::Echo(es) | StmtKind::Delete(es) => es.iter().any(|e| ey(e)),
            StmtKind::Export { declaration, default, .. } => {
                declaration.as_ref().map_or(false, |s| sy(s))
                || default.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::With { body, .. } | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. } | StmtKind::NamespaceDecl { body, .. } => body.iter().any(|s| sy(s)),
            StmtKind::MatchStatement { subject, cases } => {
                ey(subject) || cases.iter().any(|c| {
                    c.guard.as_ref().map_or(false, |e| ey(e)) || c.body.iter().any(|s| sy(s))
                })
            }
            StmtKind::Assert { test, msg } => ey(test) || msg.as_ref().map_or(false, |e| ey(e)),
            _ => false,
        }
    }
    stmts.iter().any(|s| sy(s))
}

fn walk_func_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let is_async = pair.as_rule() == Rule::async_function_declaration;
    let inner = pair.into_inner();
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::function_body => body = walk_body(p)?,
            Rule::async_kw => {}
            _ => {}
        }
    }

    let is_generator = body_contains_yield(&body);
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub: false,
    })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::param)
        .map(walk_param)
        .collect()
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let src = pair.as_str();
    let is_rest = src.starts_with("...");
    let mut name = String::new();
    let mut default = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            _ => default = Some(walk_expression(p)?),
        }
    }
    Ok(Param {
        name,
        type_hint: None,
        default,
        pass_by: PassBy::Value,
        is_rest,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .map(walk_statement)
        .collect()
}

// ── Class declaration ───────────────────────────────────────────────────────

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::assignment_expression | Rule::conditional_expression |
            Rule::logical_expr | Rule::comparison => {
                // extends expression — extract name
                parents.push(extract_ident_name(&p));
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        members.push(walk_class_member(m)?);
                    }
                }
            }
            _ => {}
        }
    }

    // Extract static block bodies — these run immediately after class definition.
    // Collect them and emit as post-class statements in a wrapping Block.
    let mut static_init_stmts: Vec<Statement> = Vec::new();
    members.retain(|m| {
        if let ClassMember::Method(ref func) = m {
            if let StmtKind::FunctionDecl { name: ref mname, ref body, ref modifiers, .. } = func.kind {
                if mname == "__static_init" && modifiers.is_static {
                    static_init_stmts.extend(body.iter().cloned());
                    return false; // remove from members
                }
            }
        }
        true
    });

    let class_stmt = StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
    };

    if static_init_stmts.is_empty() {
        Ok(class_stmt)
    } else {
        // Wrap: class declaration + static init statements in a Block
        let mut block = vec![Statement::new(class_stmt)];
        block.extend(static_init_stmts);
        Ok(StmtKind::Block(block))
    }
}

fn walk_class_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut inner_pairs: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for static keyword
    if inner_pairs.first().map_or(false, |p| p.as_rule() == Rule::static_kw) {
        is_static = true;
        inner_pairs.remove(0);
    }

    let member_pair = inner_pairs.into_iter().next()
        .ok_or("Empty class member")?;

    // ES2022 static block — convert to a synthetic static method __static_init
    if member_pair.as_rule() == Rule::static_block {
        let stmts: Vec<Statement> = member_pair.into_inner()
            .filter(|p| !matches!(p.as_rule(), Rule::NEWLINE | Rule::static_kw))
            .map(walk_statement)
            .collect::<Result<_, _>>()?;
        let func = Statement::new(StmtKind::FunctionDecl {
            name: "__static_init".to_string(),
            params: vec![],
            return_type: None,
            body: stmts,
            modifiers: Modifiers { is_static: true, ..Default::default() },
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: true,
        });
        return Ok(ClassMember::Method(Box::new(func)));
    }

    match member_pair.as_rule() {
        Rule::getter_method => {
            let mut name = String::new();
            let mut body = Vec::new();
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            Ok(ClassMember::Property {
                name,
                type_hint: None,
                getter: Some(body),
                setter: None,
                is_auto: false,
                modifiers: Modifiers { is_static, ..Default::default() },
            })
        }
        Rule::setter_method => {
            let mut name = String::new();
            let mut param = Param { name: "value".into(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false };
            let mut body = Vec::new();
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                    Rule::param => param = walk_param(p)?,
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            Ok(ClassMember::Property {
                name,
                type_hint: None,
                getter: None,
                setter: Some(PropertySetter { param, body }),
                is_auto: false,
                modifiers: Modifiers { is_static, ..Default::default() },
            })
        }
        Rule::class_method => {
            let mut name = String::new();
            let mut params = Vec::new();
            let mut body = Vec::new();
            let mut is_async = false;
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                    Rule::param_list => params = walk_params(p)?,
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            if name == "constructor" {
                Ok(ClassMember::Constructor {
                    params,
                    body,
                    base_args: None,
                    visibility: Visibility::Public,
                })
            } else {
                let is_generator = body_contains_yield(&body);
                Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                    name,
                    params,
                    return_type: None,
                    body,
                    modifiers: Modifiers { is_static, ..Default::default() },
                    handles: Vec::new(),
                    is_async,
                    is_generator,
                    is_sub: false,
                }))))
            }
        }
        Rule::class_property => {
            let mut name = String::new();
            let mut init = None;
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                    _ => init = Some(walk_expression(p)?),
                }
            }
            Ok(ClassMember::Field {
                name,
                type_hint: None,
                init,
                modifiers: Modifiers { is_static, ..Default::default() },
                with_events: false,
                array_bounds: None,
            })
        }
        other => Err(format!("Unexpected class member: {:?}", other)),
    }
}

// ── Control flow ────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let then_stmt = walk_statement(next_meaningful(&mut inner)?)?;
    // Skip NEWLINEs to find the optional else clause. The grammar's
    // eat_terminators between then and else may leave visible NEWLINE
    // tokens as siblings.
    let else_body = match next_meaningful(&mut inner) {
        Ok(p) => Some(vec![walk_statement(p)?]),
        Err(_) => None,
    };
    Ok(StmtKind::If {
        cond,
        then_body: vec![then_stmt],
        elifs: Vec::new(),
        else_body,
    })
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let header = next_rule(&mut inner, Rule::for_header)
        .or_else(|_| next_meaningful(&mut inner))?;
    let header_inner = header.into_inner().next().ok_or("Empty for header")?;
    let body_pair = next_meaningful(&mut inner)?;
    let body = vec![walk_statement(body_pair)?];

    match header_inner.as_rule() {
        Rule::for_in_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let (var, prefix) = extract_for_target(&parts)?;
            let iter = walk_expression(parts.into_iter()
                .filter(|p| !matches!(p.as_rule(), Rule::var_kind | Rule::ident_name | Rule::binding_pattern))
                .next().ok_or("missing iter expr")?)?;
            let mut full_body = prefix;
            full_body.extend(body);
            Ok(StmtKind::ForIn {
                var, key: None, iter, body: full_body, of: false,
                else_body: None, is_async: false,
            })
        }
        Rule::for_of_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let (var, prefix) = extract_for_target(&parts)?;
            let iter = walk_expression(parts.into_iter()
                .filter(|p| !matches!(p.as_rule(), Rule::var_kind | Rule::ident_name | Rule::binding_pattern))
                .next().ok_or("missing iter expr")?)?;
            let mut full_body = prefix;
            full_body.extend(body);
            Ok(StmtKind::ForIn {
                var, key: None, iter, body: full_body, of: true,
                else_body: None, is_async: false,
            })
        }
        Rule::for_c_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let mut init = None;
            let mut cond = None;
            let mut update = None;
            let mut let_vars: Vec<String> = Vec::new(); // track `let` loop vars

            for p in parts {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("Empty for init")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let mut vi = inner.into_inner();
                                let kind_pair = next_rule(&mut vi, Rule::var_kind)?;
                                let var_kind = match kind_pair.as_str() {
                                    "var" => VarDeclKind::Var,
                                    "let" => VarDeclKind::Let,
                                    "const" => VarDeclKind::Const,
                                    _ => VarDeclKind::Let,
                                };
                                let mut decls = Vec::new();
                                for d in vi {
                                    if d.as_rule() == Rule::var_declarator {
                                        let decl = walk_var_declarator(d)?;
                                        if var_kind == VarDeclKind::Let || var_kind == VarDeclKind::Const {
                                            if let BindingPattern::Ident(ref name) = decl.pattern {
                                                let_vars.push(name.clone());
                                            }
                                        }
                                        decls.push(decl);
                                    }
                                }
                                init = Some(Box::new(Statement::new(StmtKind::VarDecl {
                                    declarations: decls, kind: var_kind,
                                })));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        let expr = walk_expression(p)?;
                        if cond.is_none() && init.is_some() {
                            cond = Some(expr);
                        } else if init.is_none() {
                            // First expression could be init if no var decl
                            init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                        } else {
                            update = Some(expr);
                        }
                    }
                    _ => {
                        // Try as expression
                        if let Ok(expr) = walk_expression(p) {
                            if cond.is_none() { cond = Some(expr); }
                            else { update = Some(expr); }
                        }
                    }
                }
            }

            // Per-iteration `let` binding: wrap body in IIFE so closures
            // capture a fresh copy each iteration. Only apply when the body
            // contains function expressions/arrows that could close over the
            // loop variable — otherwise IIFE breaks break/continue.
            let body = if !let_vars.is_empty() && body_contains_closure(&body, &let_vars) {
                let params: Vec<Param> = let_vars.iter().map(|v| Param {
                    name: v.clone(), type_hint: None, default: None,
                    pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
                    is_optional: false, is_nullable: false,
                }).collect();
                let args: Vec<Argument> = let_vars.iter().map(|v| {
                    Argument::positional(Expression::ident(v))
                }).collect();
                let iife = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(body),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    args,
                    optional: false,
                });
                vec![Statement::new(StmtKind::Expr(iife))]
            } else {
                body
            };

            Ok(StmtKind::For { init, cond, update, body })
        }
        other => Err(format!("Unexpected for header: {:?}", other)),
    }
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    Ok(StmtKind::While { cond, body, else_body: None })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::DoWhile { body, cond, until: false })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(next_meaningful(&mut inner)?)?;
    let mut cases = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::switch_case {
            // Grammar: switch_case = { ("case" expression | "default") ~ ":" ~ statements }
            // Detect default by looking at the source slice.
            // Default is emitted as a SwitchCase with empty conditions,
            // preserving its position among the other cases for fallthrough.
            let is_default = p.as_str().trim_start().starts_with("default");
            let mut case_inner = p.into_inner();
            if is_default {
                let stmts: Vec<Statement> = case_inner
                    .filter(|p| p.as_rule() != Rule::NEWLINE)
                    .map(walk_statement)
                    .collect::<Result<Vec<_>, _>>()?;
                cases.push(SwitchCase {
                    conditions: vec![], // empty = default
                    body: stmts,
                });
            } else {
                let first = case_inner.next().ok_or("Empty switch case")?;
                let val = walk_expression(first)?;
                let stmts: Vec<Statement> = case_inner
                    .filter(|p| p.as_rule() != Rule::NEWLINE)
                    .map(walk_statement)
                    .collect::<Result<Vec<_>, _>>()?;
                cases.push(SwitchCase {
                    conditions: vec![CaseCondition::Value(val)],
                    body: stmts,
                });
            }
        }
    }
    Ok(StmtKind::Switch { expr, cases, default: None })
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .find(|p| p.as_rule() != Rule::NEWLINE)
        .map(walk_expression)
        .transpose()?;
    Ok(StmtKind::Return(expr))
}

fn walk_break(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let label = pair.into_inner()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string());
    Ok(StmtKind::Break(match label {
        Some(l) => BreakTarget::Label(l),
        None => BreakTarget::Implicit,
    }))
}

fn walk_continue(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let label = pair.into_inner()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string());
    Ok(StmtKind::Continue(match label {
        Some(l) => ContinueTarget::Label(l),
        None => ContinueTarget::Implicit,
    }))
}

fn walk_throw(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = walk_expression(first_meaningful(pair)?)?;
    Ok(StmtKind::Throw { expr: Some(expr), cause: None })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => body = walk_body_from_block(p)?,
            Rule::catch_clause => {
                let mut var_name = None;
                let mut catch_body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::ident_name => var_name = Some(cp.as_str().to_string()),
                        Rule::block_statement => catch_body = walk_body_from_block(cp)?,
                        _ => {}
                    }
                }
                catches.push(CatchClause {
                    types: Vec::new(),
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_body_from_block(fp)?);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::Try { body, catches, else_body: None, finally })
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let label = next_meaningful(&mut inner)?.as_str().to_string();
    let body = walk_statement(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::Labeled { label, body: Box::new(body) })
}

// ── Import / Export ─────────────────────────────────────────────────────────

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut source = String::new();
    let mut names = Vec::new();
    let mut default_name = None;
    let mut namespace_name = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::string_literal => source = unquote(p.as_str()),
            Rule::import_clause => {
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::default_import => {
                            default_name = Some(cp.as_str().to_string());
                        }
                        Rule::namespace_import => {
                            for np in cp.into_inner() {
                                if np.as_rule() == Rule::ident_name {
                                    namespace_name = Some(np.as_str().to_string());
                                }
                            }
                        }
                        Rule::named_imports => {
                            for sp in cp.into_inner() {
                                if sp.as_rule() == Rule::import_specifier {
                                    let mut parts = sp.into_inner();
                                    let name = parts.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                                    let alias = parts.next().map(|p| p.as_str().to_string());
                                    names.push(ImportName { name, alias });
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let kind = if let Some(ns) = namespace_name {
        ImportKind::Wildcard { path: source, alias: Some(ns) }
    } else if let Some(def) = default_name {
        if names.is_empty() {
            ImportKind::Default { path: source, local: def }
        } else {
            // import default, { named } from "mod" — use Named with default as first
            let mut all_names = vec![ImportName { name: "default".into(), alias: Some(def) }];
            all_names.extend(names);
            ImportKind::Named { path: source, names: all_names, level: 0 }
        }
    } else if !names.is_empty() {
        ImportKind::Named { path: source, names, level: 0 }
    } else {
        ImportKind::Simple { path: source, alias: None }
    };

    Ok(Import { kind, span })
}

fn walk_export(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut declaration = None;
    let mut names = Vec::new();
    let mut default_expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::function_declaration | Rule::async_function_declaration |
            Rule::class_declaration | Rule::variable_declaration => {
                declaration = Some(Box::new(walk_statement(p)?));
            }
            Rule::export_specifier => {
                let mut parts = p.into_inner();
                let name = parts.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                let alias = parts.next().map(|p| p.as_str().to_string());
                names.push(ExportName { name, alias });
            }
            _ => {
                // default expression
                if let Ok(expr) = walk_expression(p) {
                    default_expr = Some(Box::new(expr));
                }
            }
        }
    }

    Ok(StmtKind::Export { declaration, names, default: default_expr })
}

// ── Expressions ─────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // Literals
        Rule::numeric_literal => {
            // ES2021 numeric separator: strip `_` from digits before parsing
            let raw = pair.as_str();
            let s_owned: String = raw.chars().filter(|c| *c != '_').collect();
            let s = s_owned.as_str();
            if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(s.parse().map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0o") || s.starts_with("0O") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 8).map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0b") || s.starts_with("0B") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?)))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(unquote(pair.as_str())))),
        Rule::regex_literal => Ok(ExprKind::Lit(Literal::Str(pair.as_str().to_string()))),
        Rule::ident_name | Rule::ident_or_keyword => {
            let name = pair.as_str();
            match name {
                "true" => Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => Ok(ExprKind::Lit(Literal::Null)),
                "undefined" => Ok(ExprKind::Lit(Literal::Undefined)),
                "this" => Ok(ExprKind::This),
                "super" => Ok(ExprKind::Super),
                _ => Ok(ExprKind::Ident(name.to_string())),
            }
        }
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::undefined_kw => Ok(ExprKind::Lit(Literal::Undefined)),
        Rule::this_kw => Ok(ExprKind::This),
        Rule::super_kw => Ok(ExprKind::Super),

        // Sequence (comma expression)
        Rule::expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner()
                .filter(|p| p.as_rule() != Rule::NEWLINE)
                .collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let exprs: Vec<Expression> = inner.into_iter()
                    .map(walk_expression)
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(ExprKind::Sequence(exprs))
            }
        }

        // Assignment
        Rule::assignment_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() == 3 {
                let left = walk_expression(inner.remove(0))?;
                let op_str = inner.remove(0).as_str();
                let right = walk_expression(inner.remove(0))?;
                if op_str == "=" {
                    Ok(ExprKind::Assign { target: Box::new(left), value: Box::new(right) })
                } else {
                    // Compound assign — but this is expression level, wrap as assign
                    let op = match op_str {
                        "+=" => CompoundOp::Add, "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul, "/=" => CompoundOp::Div,
                        "%=" => CompoundOp::Mod, "**=" => CompoundOp::Pow,
                        "&=" => CompoundOp::BitAnd, "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor, "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr, ">>>=" => CompoundOp::UShr,
                        "&&=" => CompoundOp::And, "||=" => CompoundOp::Or,
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

        // Ternary
        Rule::conditional_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() == 3 {
                let cond = walk_expression(inner.remove(0))?;
                let then = walk_expression(inner.remove(0))?;
                let else_ = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Ternary { cond: Box::new(cond), then: Box::new(then), else_: Box::new(else_) })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // Binary chains
        Rule::logical_expr | Rule::comparison |
        Rule::additive | Rule::multiplicative => {
            walk_binary_chain(pair)
        }

        // Unary
        Rule::unary => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty unary")?;
            // If it's a postfix (no unary_op), delegate
            if first.as_rule() == Rule::postfix {
                return walk_expr_kind(first);
            }
            // unary_op ~ unary
            let op_str = first.as_str().trim();
            let operand = walk_expression(inner.next().ok_or("Missing unary operand")?)?;
            if op_str.starts_with("typeof") { return Ok(ExprKind::TypeOf(Box::new(operand))); }
            if op_str.starts_with("void") { return Ok(ExprKind::Void(Box::new(operand))); }
            if op_str.starts_with("delete") { return Ok(ExprKind::Delete(Box::new(operand))); }
            if op_str.starts_with("await") { return Ok(ExprKind::Await(Box::new(operand))); }
            let op = match op_str {
                "-" => UnaryOp::Neg, "+" => UnaryOp::Pos,
                "!" => UnaryOp::Not, "~" => UnaryOp::BitNot,
                "++" => UnaryOp::PreInc, "--" => UnaryOp::PreDec,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary { op, expr: Box::new(operand) })
        }

        // Postfix
        Rule::postfix => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let base = walk_expression(inner.remove(0))?;
            // Check for postfix_op (++/--)
            let has_postfix = inner.iter().any(|p| p.as_rule() == Rule::postfix_op);
            if !has_postfix {
                return Ok(base.kind);
            }
            let op_pair = inner.iter().find(|p| p.as_rule() == Rule::postfix_op).unwrap();
            let op = match op_pair.as_str() {
                "++" => UnaryOp::PostInc,
                "--" => UnaryOp::PostDec,
                _ => return Ok(base.kind),
            };
            Ok(ExprKind::Unary { op, expr: Box::new(base) })
        }

        // Call / member / index chain
        Rule::call_expression => walk_call_chain(pair),
        Rule::new_expression => {
            // new_expression = { "new" ~ primary ~ call_chain* }
            // Per JS spec: the FIRST `()` after `new` is the constructor args.
            // Any subsequent member/call/index chains are applied to the RESULT
            // of the construction (e.g. `new Foo().bar().baz`).
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty new")?;
            let mut expr = walk_expression(first)?;
            let chains: Vec<Pair<Rule>> = inner.filter(|p| p.as_rule() == Rule::call_chain).collect();
            let mut new_consumed = false; // True after the first `(args)` is processed

            for chain in chains {
                let chain_src = chain.as_str().trim_start();
                let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

                if !new_consumed && chain_src.starts_with("(") {
                    // First parens — these are the constructor args.
                    let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                        walk_arguments(arg_pair)?
                    } else { Vec::new() };
                    expr = Expression::new(ExprKind::New { class: Box::new(expr), args });
                    new_consumed = true;
                } else if !new_consumed && chain_src.starts_with(".") {
                    // Member access BEFORE constructor args: `new Foo.Bar(42)`.
                    let name = chain_inner.into_iter()
                        .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr), field: name, null_safe: false,
                    });
                } else {
                    // Chain AFTER `new X(...)` — applied to the constructed object.
                    // Handle: `(args)` call, `.member`, `?.member`, `?.()`, `[idx]`, tagged template.
                    if chain_src.starts_with("?.") {
                        if chain_inner.first().map_or(false, |p| p.as_rule() == Rule::argument_list || p.as_str().starts_with("(")) {
                            let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                                walk_arguments(arg_pair)?
                            } else { Vec::new() };
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr), args, optional: true,
                            });
                        } else {
                            let name = chain_inner.into_iter()
                                .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name || p.as_rule() == Rule::private_name)
                                .map(|p| p.as_str().to_string())
                                .unwrap_or_default();
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr), field: name, null_safe: true,
                            });
                        }
                    } else if chain_src.starts_with("(") {
                        let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                            walk_arguments(arg_pair)?
                        } else { Vec::new() };
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr), args, optional: false,
                        });
                    } else if chain_src.starts_with(".") {
                        let name = chain_inner.into_iter()
                            .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name || p.as_rule() == Rule::private_name)
                            .map(|p| p.as_str().to_string())
                            .unwrap_or_default();
                        expr = canonicalize_member_access(expr, &name);
                    } else if chain_src.starts_with("[") {
                        let index_expr = chain_inner.into_iter()
                            .find(|p| p.as_rule() == Rule::expression || matches!(p.as_rule(), Rule::assignment_expression | Rule::conditional_expression | Rule::ident_name | Rule::numeric_literal | Rule::string_literal))
                            .map(walk_expression)
                            .transpose()?
                            .unwrap_or(Expression::new(ExprKind::Lit(Literal::Int(0))));
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr), index: Box::new(index_expr),
                        });
                    }
                }
            }

            // If `new` had no `()` (e.g., `new X`), wrap the bare class.
            if !new_consumed {
                expr = Expression::new(ExprKind::New { class: Box::new(expr), args: Vec::new() });
            }
            Ok(expr.kind)
        }

        // Primary
        Rule::primary => {
            // Keyword literals (true/false/null/undefined/this/super) don't produce
            // inner pairs in pest — they're anonymous literals. Check as_str() first.
            let src = pair.as_str().trim();
            match src {
                "true" => return Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => return Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => return Ok(ExprKind::Lit(Literal::Null)),
                "undefined" => return Ok(ExprKind::Lit(Literal::Undefined)),
                "this" => return Ok(ExprKind::This),
                "super" => return Ok(ExprKind::Super),
                _ => {}
            }
            let inner = pair.into_inner().next().ok_or("Empty primary")?;
            walk_expr_kind(inner)
        }

        // Arrow functions
        Rule::arrow_function | Rule::async_arrow_function => {
            let is_async = pair.as_rule() == Rule::async_arrow_function;
            let mut params = Vec::new();
            let mut body = LambdaBody::Expr(Box::new(Expression::new(ExprKind::Lit(Literal::Null))));
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::ident_name => params = vec![Param {
                        name: p.as_str().to_string(), type_hint: None, default: None,
                        pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
                        is_optional: false, is_nullable: false,
                    }],
                    Rule::param_list => params = walk_params(p)?,
                    Rule::arrow_body => {
                        let inner = p.into_inner().next().ok_or("Empty arrow body")?;
                        body = match inner.as_rule() {
                            Rule::function_body => LambdaBody::Block(walk_body(inner)?),
                            _ => LambdaBody::Expr(Box::new(walk_expression(inner)?)),
                        };
                    }
                    Rule::function_body => body = LambdaBody::Block(walk_body(p)?),
                    Rule::async_kw => {}
                    _ => {
                        // Could be direct expression or function_body
                        if let Ok(stmts) = walk_body(p.clone()) {
                            body = LambdaBody::Block(stmts);
                        } else {
                            body = LambdaBody::Expr(Box::new(walk_expression(p)?));
                        }
                    }
                }
            }
            Ok(ExprKind::Lambda { params, body, is_async, captures: Vec::new() })
        }

        // Function expression
        Rule::function_expression | Rule::async_function_expression => {
            let stmt_kind = walk_func_decl(pair)?;
            Ok(ExprKind::FunctionExpr(Box::new(Statement::new(stmt_kind))))
        }

        // Class expression
        Rule::class_expression => {
            let mut name = None;
            let mut parent = None;
            let mut members = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::ident_name => name = Some(p.as_str().to_string()),
                    Rule::class_body => {
                        for m in p.into_inner() {
                            if m.as_rule() == Rule::class_member {
                                members.push(walk_class_member(m)?);
                            }
                        }
                    }
                    _ => parent = Some(Box::new(walk_expression(p)?)),
                }
            }
            Ok(ExprKind::ClassExpr { name, parent, members })
        }

        // Array literal
        Rule::array_literal => {
            let elements = pair.into_inner()
                .filter(|p| p.as_rule() == Rule::array_element)
                .map(|p| {
                    let src = p.as_str();
                    let spread = src.trim_start().starts_with("...");
                    let inner = p.into_inner().next().ok_or("Empty array element".to_string())?;
                    let value = walk_expression(inner)?;
                    Ok(ArrayElement { key: None, value, spread, by_ref: false })
                })
                .collect::<Result<Vec<_>, String>>()?;
            Ok(ExprKind::Array(elements))
        }

        // Object literal
        Rule::object_literal => {
            let props = pair.into_inner()
                .filter(|p| p.as_rule() == Rule::object_property)
                .map(walk_object_property)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(ExprKind::Object(props))
        }

        // Template literal
        Rule::template_literal => {
            let mut parts = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::template_full => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(s[1..s.len()-1].to_string()));
                    }
                    Rule::template_head => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(s[1..s.len()-2].to_string()));
                    }
                    Rule::template_middle => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(s[1..s.len()-2].to_string()));
                    }
                    Rule::template_tail => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(s[1..s.len()-1].to_string()));
                    }
                    _ => parts.push(InterpolPart::Expr(walk_expression(p)?)),
                }
            }
            Ok(ExprKind::Interpolation(parts))
        }

        // Spread
        Rule::argument => {
            let src = pair.as_str();
            let spread = src.trim_start().starts_with("...");
            let inner = pair.into_inner().next().ok_or("Empty argument")?;
            let expr = walk_expression(inner)?;
            if spread {
                Ok(ExprKind::Spread(Box::new(expr)))
            } else {
                Ok(expr.kind)
            }
        }

        // Passthrough wrappers
        Rule::call_chain | Rule::property_name | Rule::computed_property_name => {
            let inner = pair.into_inner().next().ok_or("Empty wrapper")?;
            walk_expr_kind(inner)
        }

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Binary chain walker ─────────────────────────────────────────────────────

fn walk_binary_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let _rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    // First operand
    let mut left = walk_expression(inner.remove(0))?;

    // Remaining: (op, operand) pairs
    let mut i = 0;
    while i + 1 < inner.len() {
        let op_pair = &inner[i];
        let op = match op_pair.as_rule() {
            Rule::logical_op | Rule::nullish_op | Rule::or_op | Rule::and_op |
            Rule::bitor_op | Rule::bitxor_op | Rule::bitand_op |
            Rule::comparison_op | Rule::equality_op | Rule::relational_op | Rule::shift_op |
            Rule::additive_op | Rule::multiplicative_op | Rule::mul_op | Rule::exp_op => op_pair.as_str().trim(),
            _ => op_pair.as_str().trim(),
        };
        let right = walk_expression(inner[i + 1].clone())?;

        let bin_op = match op {
            "??" => BinOp::NullCoalesce,
            "||" => BinOp::Or, "&&" => BinOp::And,
            "|" => BinOp::BitOr, "^" => BinOp::BitXor, "&" => BinOp::BitAnd,
            "===" => BinOp::StrictEq, "!==" => BinOp::StrictNotEq,
            "==" => BinOp::Eq, "!=" => BinOp::NotEq,
            "<" => BinOp::Lt, ">" => BinOp::Gt,
            "<=" => BinOp::LtEq, ">=" => BinOp::GtEq,
            "instanceof" => BinOp::InstanceOf, "in" => BinOp::In,
            ">>>" => BinOp::UShr, ">>" => BinOp::Shr, "<<" => BinOp::Shl,
            "+" => BinOp::Add, "-" => BinOp::Sub,
            "*" => BinOp::Mul, "/" => BinOp::Div, "%" => BinOp::Mod,
            "**" => BinOp::Pow,
            _ => BinOp::Add,
        };

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        i += 2;
    }

    Ok(left.kind)
}

// ── Call chain walker ───────────────────────────────────────────────────────

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty call expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain { continue; }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with("?.") {
            // Optional chaining
            // Detect optional call: ?.(...) — chain_inner may be empty (no args) or contain argument_list.
            // Use chain_src to detect the "(" after "?." since grammar literals aren't in chain_inner.
            let is_optional_call = chain_src.starts_with("?.(")
                || chain_inner.first().map_or(false, |p| p.as_rule() == Rule::argument_list);
            if is_optional_call {
                // ?.( call
                let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                    walk_arguments(arg_pair)?
                } else { Vec::new() };
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(expr), args, optional: true,
                });
            } else {
                // ?. member
                let name = chain_inner.into_iter()
                    .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name || p.as_rule() == Rule::private_name)
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr), field: name, null_safe: true,
                });
            }
        } else if chain_src.starts_with("(") {
            // Call
            let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                walk_arguments(arg_pair)?
            } else { Vec::new() };
            expr = Expression::new(ExprKind::Call {
                callee: Box::new(expr), args, optional: false,
            });
        } else if chain_src.starts_with(".") {
            // Member access — normalize JS .length to canonical __len__
            let name = chain_inner.into_iter()
                .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name || p.as_rule() == Rule::private_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            expr = canonicalize_member_access(expr, &name);
        } else if chain_src.starts_with("[") {
            // Computed / index
            let index_expr = chain_inner.into_iter()
                .find(|p| p.as_rule() == Rule::expression || matches!(p.as_rule(), Rule::assignment_expression | Rule::conditional_expression | Rule::ident_name | Rule::numeric_literal | Rule::string_literal))
                .map(walk_expression)
                .transpose()?
                .unwrap_or(Expression::new(ExprKind::Lit(Literal::Int(0))));
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr), index: Box::new(index_expr),
            });
        } else if chain_src.starts_with("`") {
            // Tagged template: tag`parts...${expr}...`
            // Desugar to: tag(["part0", "part1", ...], expr0, expr1, ...)
            if let Some(tmpl) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::template_literal) {
                let (parts, exprs) = walk_template_parts(tmpl)?;
                let mut args: Vec<Argument> = Vec::new();
                // First arg: array of string parts
                let strings_array = Expression::new(ExprKind::Array(
                    parts.into_iter().map(|s| {
                        ArrayElement {
                            key: None,
                            value: Expression::new(ExprKind::Lit(Literal::Str(s))),
                            spread: false,
                            by_ref: false,
                        }
                    }).collect()
                ));
                args.push(Argument::positional(strings_array));
                // Rest: expression values
                for e in exprs {
                    args.push(Argument::positional(e));
                }
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(expr), args, optional: false,
                });
            }
        }
    }

    // Normalize variadic concat: `x.concat(a, b, c)` → `x.concat(a).concat(b).concat(c)`
    // The stdlib concat function is binary (receiver + 1 arg). For variadic calls,
    // desugar into a chain of binary concat calls. Works for both strings and arrays.
    expr = desugar_variadic_concat(expr);

    Ok(expr.kind)
}

/// Check if a for-loop body contains closures (lambdas or function expressions)
/// that reference any of the given `let` variable names. Used to decide whether
/// to wrap the body in an IIFE for per-iteration binding.
fn body_contains_closure(stmts: &[Statement], _vars: &[String]) -> bool {
    // Simple heuristic: check if any lambda/function expression exists in the body.
    // A more precise check would verify the lambda references a let-var, but
    // the simple check is correct — IIFE is safe when there ARE closures, and
    // we skip it when there are none (preserving break/continue).
    fn has_closure_expr(expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => true,
            ExprKind::Call { callee, args, .. } => {
                has_closure_expr(callee) || args.iter().any(|a| has_closure_expr(&a.value))
            }
            ExprKind::Member { object, .. } => has_closure_expr(object),
            ExprKind::Binary { left, right, .. } => has_closure_expr(left) || has_closure_expr(right),
            ExprKind::Unary { expr, .. } => has_closure_expr(expr),
            ExprKind::Ternary { cond, then, else_, .. } => {
                has_closure_expr(cond) || has_closure_expr(then) || has_closure_expr(else_)
            }
            ExprKind::Array(elems) => elems.iter().any(|e| has_closure_expr(&e.value)),
            ExprKind::Index { object, index } => has_closure_expr(object) || has_closure_expr(index),
            ExprKind::Assign { target: _, value } => has_closure_expr(value),
            _ => false,
        }
    }
    fn has_closure_stmt(stmt: &Statement) -> bool {
        match &stmt.kind {
            StmtKind::Expr(e) => has_closure_expr(e),
            StmtKind::VarDecl { declarations, .. } => {
                declarations.iter().any(|d| d.init.as_ref().map_or(false, |e| has_closure_expr(e)))
            }
            StmtKind::If { cond, then_body, elifs, else_body, .. } => {
                has_closure_expr(cond) || then_body.iter().any(has_closure_stmt)
                || elifs.iter().any(|(c, b)| has_closure_expr(c) || b.iter().any(has_closure_stmt))
                || else_body.as_ref().map_or(false, |b| b.iter().any(has_closure_stmt))
            }
            StmtKind::Block(stmts) => stmts.iter().any(has_closure_stmt),
            StmtKind::Return(Some(e)) => has_closure_expr(e),
            _ => false,
        }
    }
    stmts.iter().any(has_closure_stmt)
}

/// Desugar `x.concat(a, b, c)` into `x.concat(a).concat(b).concat(c)`.
fn desugar_variadic_concat(expr: Expression) -> Expression {
    if let ExprKind::Call { ref callee, ref args, optional } = expr.kind {
        if args.len() > 1 {
            if let ExprKind::Member { ref object, ref field, null_safe } = callee.kind {
                if field == "concat" {
                    // Chain: start with receiver.concat(args[0]), then .concat(args[1]), etc.
                    let mut result = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "concat".to_string(),
                            null_safe,
                        })),
                        args: vec![args[0].clone()],
                        optional,
                    });
                    for arg in &args[1..] {
                        result = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(result),
                                field: "concat".to_string(),
                                null_safe: false,
                            })),
                            args: vec![arg.clone()],
                            optional: false,
                        });
                    }
                    return result;
                }
            }
        }
    }
    expr
}

/// Walk a template_literal into (string_parts, expressions).
/// `hello ${name}!` → (["hello ", "!"], [name])
fn walk_template_parts(pair: Pair<Rule>) -> Result<(Vec<String>, Vec<Expression>), String> {
    let mut parts: Vec<String> = Vec::new();
    let mut exprs: Vec<Expression> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::template_full => {
                let s = p.as_str();
                // Strip leading ` and trailing `
                let inner = &s[1..s.len()-1];
                parts.push(unescape_template(inner));
            }
            Rule::template_head => {
                let s = p.as_str();
                // Strip leading ` and trailing ${
                let inner = &s[1..s.len()-2];
                parts.push(unescape_template(inner));
            }
            Rule::template_middle => {
                let s = p.as_str();
                // Strip leading } and trailing ${
                let inner = &s[1..s.len()-2];
                parts.push(unescape_template(inner));
            }
            Rule::template_tail => {
                let s = p.as_str();
                // Strip leading } and trailing `
                let inner = &s[1..s.len()-1];
                parts.push(unescape_template(inner));
            }
            _ => {
                // Expression inside ${...}
                exprs.push(walk_expression(p)?);
            }
        }
    }
    Ok((parts, exprs))
}

fn unescape_template(s: &str) -> String {
    s.replace("\\`", "`").replace("\\$", "$").replace("\\\\", "\\")
}

/// Canonicalize JS property access to unified AST representation.
/// `arr.length` → `Call(__len__, [arr])`
///
/// Note: only `.length` is normalized — `.size` is too generic in JS (could be a custom property).
fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    let canonical = match name {
        "length" => Some("__len__"),
        _ => None,
    };
    if let Some(canonical_name) = canonical {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(canonical_name)),
            args: vec![Argument::positional(object)],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        })
    }
}

// JS method call canonicalization is intentionally minimal:
// Methods like .toString() may be overridden on user classes, so we leave them as
// regular method calls and let the compiler dispatch via the class method binding.
// Only true builtin operations like .length (handled in canonicalize_member_access)
// are normalized to canonical builtins.

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::argument)
        .map(|p| {
            let spread = p.as_str().trim_start().starts_with("...");
            let inner = p.into_inner().next().ok_or("Empty argument".to_string())?;
            let value = walk_expression(inner)?;
            Ok(Argument { value, name: None, by_ref: false, spread })
        })
        .collect()
}

// ── Object property walker ──────────────────────────────────────────────────

fn walk_object_property(pair: Pair<Rule>) -> Result<ObjectProperty, String> {
    let src = pair.as_str().trim();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Spread: { ...expr }
    if src.starts_with("...") {
        let expr = walk_expression(inner.remove(0))?;
        return Ok(ObjectProperty::Spread(expr));
    }

    // Computed: { [expr]: value }
    if inner.first().map_or(false, |p| p.as_rule() == Rule::computed_property_name) {
        let key_pair = inner.remove(0);
        let key = walk_expression(key_pair.into_inner().next().ok_or("Empty computed key")?)?;
        let value = walk_expression(inner.remove(0))?;
        return Ok(ObjectProperty::Computed { key, value });
    }

    // Method: { name() {} } or getter/setter
    if inner.len() >= 2 {
        let has_body = inner.iter().any(|p| p.as_rule() == Rule::function_body);
        if has_body {
            let key = inner.remove(0).as_str().to_string();
            let mut params = Vec::new();
            let mut body = Vec::new();
            for p in inner {
                match p.as_rule() {
                    Rule::param_list => params = walk_params(p)?,
                    Rule::param => params = vec![walk_param(p)?],
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            let is_generator = body_contains_yield(&body);
            let func = Statement::new(StmtKind::FunctionDecl {
                name: key.clone(), params, return_type: None, body,
                modifiers: Modifiers::default(), handles: Vec::new(),
                is_async: false, is_generator, is_sub: false,
            });
            return Ok(ObjectProperty::Method { key, value: Box::new(func) });
        }
    }

    // Key: value or shorthand
    if inner.len() == 1 {
        return Ok(ObjectProperty::Shorthand(inner.remove(0).as_str().to_string()));
    }

    if inner.len() >= 2 {
        let key_pair = inner.remove(0);
        // Object keys: identifiers become string literals (JS object keys are always strings)
        let key = match key_pair.as_rule() {
            Rule::ident_name | Rule::ident_or_keyword | Rule::property_name => {
                let key_str = key_pair.as_str().to_string();
                // property_name may contain inner pairs (string/number/ident) — extract
                if let Some(inner_pair) = key_pair.into_inner().next() {
                    match inner_pair.as_rule() {
                        Rule::string_literal => walk_expression(inner_pair)?,
                        Rule::numeric_literal => walk_expression(inner_pair)?,
                        _ => Expression::string(&key_str),
                    }
                } else {
                    Expression::string(&key_str)
                }
            }
            _ => walk_expression(key_pair)?,
        };
        let value = walk_expression(inner.remove(0))?;
        return Ok(ObjectProperty::KeyValue { key, value });
    }

    Err("Could not parse object property".into())
}

// ── Helpers ─────────────────────────────────────────────────────────────────

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

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    pair.into_inner()
        .find(|p| p.as_rule() != Rule::NEWLINE)
        .ok_or_else(|| "Expected inner pair".into())
}

fn next_meaningful<'a>(pairs: &mut impl Iterator<Item = Pair<'a, Rule>>) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if p.as_rule() != Rule::NEWLINE { return Ok(p); }
    }
    Err("Expected next pair".into())
}

fn next_rule<'a>(pairs: &mut impl Iterator<Item = Pair<'a, Rule>>, rule: Rule) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if p.as_rule() == rule { return Ok(p); }
    }
    Err(format!("Expected {:?}", rule))
}

fn walk_body_from_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .map(walk_statement)
        .collect()
}

fn extract_ident_name(pair: &Pair<Rule>) -> String {
    pair.as_str().trim().to_string()
}

fn extract_ident_from_for_target<'a>(pairs: &mut impl Iterator<Item = Pair<'a, Rule>>) -> Result<String, String> {
    for p in pairs {
        match p.as_rule() {
            Rule::ident_name => return Ok(p.as_str().to_string()),
            Rule::var_kind => continue,
            _ => continue,
        }
    }
    Err("Expected identifier in for target".into())
}

/// Extract the loop variable and any destructuring prefix statements.
/// For `for (let x of arr)` returns ("x", []).
/// For `for (let [a, b] of arr)` returns ("__forof_tmp", [VarDecl let [a,b] = __forof_tmp])
fn extract_for_target(parts: &[Pair<Rule>]) -> Result<(String, Vec<Statement>), String> {
    let mut var_kind = VarDeclKind::Let;
    for p in parts {
        match p.as_rule() {
            Rule::var_kind => {
                var_kind = match p.as_str() {
                    "var" => VarDeclKind::Var,
                    "const" => VarDeclKind::Const,
                    _ => VarDeclKind::Let,
                };
            }
            Rule::ident_name => {
                return Ok((p.as_str().to_string(), Vec::new()));
            }
            Rule::binding_pattern => {
                let inner = p.clone().into_inner().next().ok_or("Empty binding pattern")?;
                if inner.as_rule() == Rule::ident_name {
                    return Ok((inner.as_str().to_string(), Vec::new()));
                }
                // Destructuring pattern — desugar to: let __forof_tmp; let [...] = __forof_tmp
                let pattern = walk_binding_pattern(p.clone())?;
                let tmp = "__forof_tmp".to_string();
                let prefix = Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern,
                        type_hint: None,
                        init: Some(Expression::ident(&tmp)),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: var_kind,
                });
                return Ok((tmp, vec![prefix]));
            }
            _ => continue,
        }
    }
    Err("Expected identifier or binding pattern in for target".into())
}

fn unquote(s: &str) -> String {
    if s.len() < 2 { return s.to_string(); }
    let inner = &s[1..s.len()-1];
    inner.replace("\\'", "'").replace("\\\"", "\"").replace("\\\\", "\\")
        .replace("\\n", "\n").replace("\\t", "\t").replace("\\r", "\r")
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add, CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul, CompoundOp::Div => BinOp::Div,
        CompoundOp::Mod => BinOp::Mod, CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd, CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor, CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr, CompoundOp::UShr => BinOp::UShr,
        CompoundOp::And => BinOp::And, CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::IDiv => BinOp::IDiv, CompoundOp::Concat => BinOp::Concat,
    }
}
