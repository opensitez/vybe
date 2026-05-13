use pest::Parser;
use pest::iterators::Pair;
use super::{RubyParser, Rule};
use crate::ast::*;

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = RubyParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => { walk_stmt_into(top, &mut body, &mut imports)?; continue; }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                _ => walk_stmt_into(pair, &mut body, &mut imports)?,
            }
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Ruby,
        body,
        imports,
    })
}

fn walk_stmt_into(pair: Pair<Rule>, body: &mut Vec<Statement>, imports: &mut Vec<Import>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::require_stmt => imports.push(walk_require(pair)?),
        _ => {
            let stmt = walk_statement(pair)?;
            if !matches!(stmt.kind, StmtKind::Empty) {
                body.push(stmt);
            }
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::method_def => walk_method_def(pair)?,
        Rule::class_def => walk_class_def(pair)?,
        Rule::module_def => walk_module_def(pair)?,

        Rule::if_stmt => walk_if(pair)?,
        Rule::unless_stmt => walk_unless(pair)?,
        Rule::while_stmt => walk_while(pair)?,
        Rule::until_stmt => walk_until(pair)?,
        Rule::for_stmt => walk_for(pair)?,
        Rule::case_stmt => walk_case(pair)?,
        Rule::begin_stmt => walk_begin(pair)?,
        Rule::loop_stmt => walk_loop(pair)?,

        Rule::return_stmt => walk_return(pair)?,
        Rule::break_stmt => walk_break_or_next(pair, true)?,
        Rule::next_stmt => walk_break_or_next(pair, false)?,
        Rule::raise_stmt => walk_raise(pair)?,
        Rule::retry_stmt => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::redo_stmt => StmtKind::Continue(ContinueTarget::Implicit),

        Rule::require_stmt => return Ok(Statement::new(StmtKind::Empty)), // handled in walk_stmt_into
        Rule::at_exit_stmt => StmtKind::Empty, // no runtime equivalent
        Rule::catch_throw_stmt => StmtKind::Empty, // simplified
        Rule::access_modifier_stmt => StmtKind::Empty, // metadata only
        Rule::alias_stmt => StmtKind::Empty, // not directly representable
        Rule::undef_stmt => StmtKind::Empty, // not directly representable

        Rule::multi_assign_stmt => walk_multi_assign(pair)?,
        Rule::expr_or_assign_stmt => walk_expr_or_assign(pair)?,

        Rule::NEWLINE => StmtKind::Empty,

        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

// ── Method def ──────────────────────────────────────────────────────────────

fn walk_method_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut is_self_method = false;
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_name => {
                let text = p.as_str();
                if text.starts_with("self.") {
                    is_self_method = true;
                    name = text[5..].to_string();
                } else {
                    name = text.to_string();
                }
            }
            Rule::method_params => params = walk_method_params(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {}
        }
    }

    // Don't apply implicit return to constructors — the compiler handles constructor return
    if name != "initialize" {
        apply_implicit_return(&mut body);
    }

    let mut modifiers = Modifiers::default();
    if is_self_method {
        modifiers.is_static = true;
    }

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers,
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

fn walk_method_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_list {
            params = walk_param_list(p)?;
        }
    }
    Ok(params)
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_item {
            let inner = p.into_inner().next();
            if let Some(item) = inner {
                match item.as_rule() {
                    Rule::normal_param => {
                        params.push(Param {
                            name: item.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::optional_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ => default = Some(walk_expression(c)?),
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: true,
                            is_nullable: false,
                        });
                    }
                    Rule::splat_param => {
                        let name = item.into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: true,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::double_splat_param => {
                        let name = item.into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: true,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::block_param => {
                        // &block — ignore for now, blocks are handled differently
                    }
                    Rule::keyword_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ if is_expression_rule(c.as_rule()) => {
                                    default = Some(walk_expression(c)?);
                                }
                                _ => {}
                            }
                        }
                        let is_optional = default.is_some();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional,
                            is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
    }
    Ok(params)
}

// ── Class def ───────────────────────────────────────────────────────────────

fn walk_class_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::constant_path => {
                parents.push(p.as_str().to_string());
            }
            Rule::class_body => {
                members = walk_class_body(p)?;
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_class_body(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    let mut current_visibility = Visibility::Public;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::access_modifier_stmt => {
                let text = p.as_str().trim();
                if text.starts_with("private") {
                    current_visibility = Visibility::Private;
                } else if text.starts_with("protected") {
                    current_visibility = Visibility::Protected;
                } else {
                    current_visibility = Visibility::Public;
                }
            }
            Rule::attr_decl => {
                members.extend(walk_attr_decl(p)?);
            }
            Rule::method_def => {
                let stmt_kind = walk_method_def(p)?;
                if let StmtKind::FunctionDecl { name, params, body, modifiers, .. } = &stmt_kind {
                    if name == "initialize" {
                        // Extract instance variable assignments from constructor body
                        members.push(ClassMember::Constructor {
                            params: params.clone(),
                            body: body.clone(),
                            base_args: None,
                            visibility: current_visibility,
                        });
                    } else {
                        let mut mods = modifiers.clone();
                        mods.visibility = current_visibility;
                        members.push(ClassMember::Method(Box::new(
                            Statement::new(stmt_kind)
                        )));
                    }
                }
            }
            Rule::include_stmt | Rule::extend_stmt => {
                // Include/extend — ignore for now
            }
            Rule::alias_stmt => {}
            Rule::class_def => {
                // Nested class
                let nested = walk_class_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::module_def => {
                let nested = walk_module_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::NEWLINE => {}
            _ => {
                // Other statements in class body → treat as static initializer
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    members.push(ClassMember::Method(Box::new(stmt)));
                }
            }
        }
    }
    Ok(members)
}

fn walk_attr_decl(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut kind = "";
    let mut names = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::attr_kind => kind = match p.as_str().trim() {
                "attr_accessor" => "accessor",
                "attr_reader" => "reader",
                "attr_writer" => "writer",
                _ => "accessor",
            },
            Rule::symbol_list => {
                for s in p.into_inner() {
                    let text = s.as_str().trim();
                    let name = if text.starts_with(':') {
                        &text[1..]
                    } else if text.starts_with('"') || text.starts_with('\'') {
                        &text[1..text.len()-1]
                    } else {
                        text
                    };
                    names.push(name.to_string());
                }
            }
            _ => {}
        }
    }

    let mut members = Vec::new();
    for name in names {
        let has_getter = kind == "accessor" || kind == "reader";
        let has_setter = kind == "accessor" || kind == "writer";

        // Getter → method `name()` that returns self._rb_<field>
        // The backing field is created by `@name = ...` in initialize
        // which maps to self._rb_name (prefixed to avoid struct key collision)
        if has_getter {
            let self_expr = Expression::new(ExprKind::Ident("self".into()));
            let field_access = Expression::new(ExprKind::Member {
                object: Box::new(self_expr),
                field: format!("_rb_{}", name),
                null_safe: false,
            });
            let body = vec![Statement::new(StmtKind::Return(Some(field_access)))];
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: name.clone(),
                    params: Vec::new(),
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        }

        // Setter semantics: Ruby `d.name = x` is transformed in the walker to
        // Assign(Member(d, "_rb_name"), x) via fixup_assign_target, which writes
        // directly to the _rb_ prefixed backing field via struct_set.
        let _ = has_setter;
    }
    Ok(members)
}

// ── Module def ──────────────────────────────────────────────────────────────

fn walk_module_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::class_body => {
                members = walk_class_body(p)?;
            }
            _ => {}
        }
    }

    Ok(StmtKind::ModuleDecl {
        name,
        members,
        visibility: Visibility::Public,
    })
}

// ── If ──────────────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression if_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::if_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier if body")?)?;
        // skip if_kw
        iter.find(|p| p.as_rule() == Rule::if_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier if condition")?)?;
        return Ok(StmtKind::If {
            cond,
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form: if cond then_kw? body elsif* else? end
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(next_rule(&mut iter, Rule::body)?)?;

    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in iter {
        match p.as_rule() {
            Rule::elsif_clause => {
                let mut ei = p.into_inner();
                let econd = walk_expression(next_meaningful(&mut ei)?)?;
                let ebody = walk_body(find_rule(ei, Rule::body)?)?;
                elifs.push((econd, ebody));
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If { cond, then_body, elifs, else_body })
}

// ── Unless ──────────────────────────────────────────────────────────────────

fn walk_unless(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression unless_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::unless_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier unless body")?)?;
        iter.find(|p| p.as_rule() == Rule::unless_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier unless condition")?)?;
        // unless → if !cond
        return Ok(StmtKind::If {
            cond: negate(cond),
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;

    let mut else_body = None;
    for p in iter {
        if p.as_rule() == Rule::else_clause {
            let ei = p.into_inner();
            else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
        }
    }

    // unless cond → if !cond
    Ok(StmtKind::If {
        cond: negate(cond),
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── While ───────────────────────────────────────────────────────────────────

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form: expression while_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::while_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier while body")?)?;
        iter.find(|p| p.as_rule() == Rule::while_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier while condition")?)?;
        return Ok(StmtKind::While {
            cond,
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While { cond, body, else_body: None })
}

// ── Until ───────────────────────────────────────────────────────────────────

fn walk_until(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form
    if children.iter().any(|p| p.as_rule() == Rule::until_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier until body")?)?;
        iter.find(|p| p.as_rule() == Rule::until_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier until condition")?)?;
        // until cond → while !cond
        return Ok(StmtKind::While {
            cond: negate(cond),
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form: until cond → while !cond
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While { cond: negate(cond), body, else_body: None })
}

// ── For ─────────────────────────────────────────────────────────────────────

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut vars = Vec::new();
    let mut iter_expr = None;
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => vars.push(p.as_str().to_string()),
            Rule::in_kw | Rule::do_kw => {}
            Rule::body => body = walk_body(p)?,
            _ if is_expression_rule(p.as_rule()) => {
                if iter_expr.is_none() {
                    iter_expr = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    // Multi-target destructuring
    let var = if vars.len() > 1 {
        let tmp = "__forin_element".to_string();
        let mut destructure_stmts: Vec<Statement> = Vec::new();
        for (i, name) in vars.iter().enumerate() {
            destructure_stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    index: Box::new(Expression::int(i as i64)),
                    null_safe: false,
                }),
            }));
        }
        destructure_stmts.extend(body);
        body = destructure_stmts;
        tmp
    } else {
        vars.into_iter().next().unwrap_or_default()
    };

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter: iter_expr.unwrap_or(Expression::null()),
        body,
        of: true,
        else_body: None,
        is_async: false,
    })
}

// ── Case / When ─────────────────────────────────────────────────────────────

fn walk_case(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut subject = None;
    let mut cases = Vec::new();
    let mut default = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::when_clause => {
                let mut conditions = Vec::new();
                let mut body = Vec::new();
                for wp in p.into_inner() {
                    match wp.as_rule() {
                        Rule::expression_list => {
                            for ep in wp.into_inner() {
                                if is_expression_rule(ep.as_rule()) {
                                    let expr = walk_expression(ep)?;
                                    conditions.push(CaseCondition::Value(expr));
                                }
                            }
                        }
                        Rule::body => body = walk_body(wp)?,
                        Rule::then_kw => {}
                        _ if is_expression_rule(wp.as_rule()) => {
                            let expr = walk_expression(wp)?;
                            conditions.push(CaseCondition::Value(expr));
                        }
                        _ => {}
                    }
                }
                cases.push(SwitchCase { conditions, body });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                default = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ if is_expression_rule(p.as_rule()) => {
                if subject.is_none() {
                    subject = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: subject.unwrap_or(Expression::bool(true)),
        cases,
        default,
    })
}

// ── Begin / Rescue / Ensure ─────────────────────────────────────────────────

fn walk_begin(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut else_body = None;
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::body => {
                if body.is_empty() {
                    body = walk_body(p)?;
                }
            }
            Rule::rescue_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();

                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::constant_path => types.push(cp.as_str().to_string()),
                        Rule::identifier => var_name = Some(cp.as_str().to_string()),
                        Rule::body => catch_body = walk_body(cp)?,
                        _ => {}
                    }
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            Rule::ensure_clause => {
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::body {
                        finally = Some(walk_body(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try { body, catches, else_body, finally })
}

// ── Loop ────────────────────────────────────────────────────────────────────

fn walk_loop(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block => {
                for dp in p.into_inner() {
                    if dp.as_rule() == Rule::body {
                        body = walk_body(dp)?;
                    }
                }
            }
            _ => {}
        }
    }
    // loop { ... } → while true { ... }
    Ok(StmtKind::While {
        cond: Expression::bool(true),
        body,
        else_body: None,
    })
}

// ── Return ──────────────────────────────────────────────────────────────────

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            exprs.push(walk_expression(p)?);
        } else if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    exprs.push(walk_expression(ep)?);
                }
            }
        }
    }
    // Ruby `return a, b` semantically returns an Array, but we model
    // it as `ExprKind::Tuple` so the compiler's multi-value pre-scan
    // can recognise the uniform-arity pattern. Tuple and Array lower
    // to the same `ecma:array` packed representation — the AST
    // distinction is purely to drive the multi-value opt-in.
    let expr = if exprs.len() > 1 {
        Some(Expression::new(ExprKind::Tuple(exprs)))
    } else {
        exprs.into_iter().next()
    };
    Ok(StmtKind::Return(expr))
}

// ── Raise ───────────────────────────────────────────────────────────────────

fn walk_raise(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        } else if is_expression_rule(p.as_rule()) && expr.is_none() {
            expr = Some(walk_expression(p)?);
        }
    }
    let stmt = StmtKind::Throw { expr, cause: None };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

// ── Break / Next with optional modifier ─────────────────────────────────────

fn walk_break_or_next(pair: Pair<Rule>, is_break: bool) -> Result<StmtKind, String> {
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        }
    }
    let stmt = if is_break {
        StmtKind::Break(BreakTarget::Implicit)
    } else {
        StmtKind::Continue(ContinueTarget::Implicit)
    };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

// ── Multi-assign ────────────────────────────────────────────────────────────

fn walk_multi_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut values = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::target => {
                let inner: Vec<Pair<Rule>> = p.into_inner().collect();
                if let Some(first) = inner.into_iter().next() {
                    targets.push(walk_expression(first)?);
                }
            }
            Rule::expression_list => {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    // Multi-assign: a, b = 1, 2
    // Emit as destructuring assign
    if values.len() == 1 {
        // a, b = [1, 2] — single RHS
        let patterns = targets.iter().map(|t| {
            if let ExprKind::Ident(name) = &t.kind {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
            } else {
                ArrayPatternElem::Hole
            }
        }).collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)))],
            value: values.into_iter().next().unwrap(),
        })
    } else {
        // a, b = 1, 2 — wrap RHS in array
        let value = Expression::new(ExprKind::Array(
            values.into_iter().map(|v| ArrayElement { key: None, value: v, spread: false, by_ref: false }).collect()
        ));
        let patterns = targets.iter().map(|t| {
            if let ExprKind::Ident(name) = &t.kind {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
            } else {
                ArrayPatternElem::Hole
            }
        }).collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)))],
            value,
        })
    }
}

// ── Expression or assignment ────────────────────────────────────────────────

/// Transform assignment targets: unwrap `Call(Member(obj, field), [])` → `Member(obj, "_rb_field")`.
/// In Ruby, `d.name = x` goes through a setter method which writes to the backing @name ivar.
/// Since @vars are stored with `_rb_` prefix, external assignments must write there too.
fn fixup_assign_target(expr: Expression) -> Expression {
    if let ExprKind::Call { ref callee, ref args, .. } = expr.kind {
        if args.is_empty() {
            if let ExprKind::Member { ref object, ref field, null_safe } = callee.kind {
                return Expression::new(ExprKind::Member {
                    object: object.clone(),
                    field: format!("_rb_{}", field),
                    null_safe,
                });
            }
        }
    }
    expr
}

fn walk_expr_or_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .collect();

    if inner.is_empty() {
        return Ok(StmtKind::Empty);
    }

    // ── Check for command call (postfix ~ command_args ~ block_literal? ~ modifier_suffix?)
    let has_command_args = inner.iter().any(|p| p.as_rule() == Rule::command_args);
    if has_command_args {
        return walk_command_call(inner);
    }

    // ── Check for augmented assignment
    let has_aug = inner.iter().any(|p| p.as_rule() == Rule::aug_assign_op);
    if has_aug {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let op_str = inner.remove(0).as_str().to_string();
        let value = if !inner.is_empty() && is_expression_rule(inner[0].as_rule()) {
            walk_expression(inner.remove(0))?
        } else {
            Expression::null()
        };
        let op = match op_str.as_str() {
            "+=" => CompoundOp::Add,
            "-=" => CompoundOp::Sub,
            "*=" => CompoundOp::Mul,
            "/=" => CompoundOp::Div,
            "%=" => CompoundOp::Mod,
            "**=" => CompoundOp::Pow,
            "<<=" => CompoundOp::Shl,
            ">>=" => CompoundOp::Shr,
            "|=" => CompoundOp::BitOr,
            "&=" => CompoundOp::BitAnd,
            "^=" => CompoundOp::BitXor,
            "||=" => CompoundOp::Or,
            "&&=" => CompoundOp::And,
            _ => CompoundOp::Add,
        };
        let stmt = StmtKind::CompoundAssign { target, op, value };
        return maybe_wrap_modifier(stmt, &mut inner);
    }

    // ── Check for regular assignment (expression = expression_list)
    let has_expr_list = inner.iter().any(|p| p.as_rule() == Rule::expression_list);
    if has_expr_list {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let mut values = Vec::new();
        let mut remaining = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::expression_list {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            } else if p.as_rule() == Rule::modifier_suffix {
                remaining.push(p);
            } else if is_expression_rule(p.as_rule()) {
                values.push(walk_expression(p)?);
            }
        }
        if values.is_empty() {
            let stmt = StmtKind::Expr(target);
            return maybe_wrap_modifier(stmt, &mut remaining);
        }
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Array(
                values.into_iter().map(|v| ArrayElement { key: None, value: v, spread: false, by_ref: false }).collect()
            ))
        };
        let stmt = StmtKind::Assign { targets: vec![target], value };
        return maybe_wrap_modifier(stmt, &mut remaining);
    }

    // ── Expression statement (expression ~ modifier_suffix?)
    let expr = walk_expression(inner.remove(0))?;
    let stmt = StmtKind::Expr(expr);
    maybe_wrap_modifier(stmt, &mut inner)
}

/// Handle command-style call: postfix ~ command_args ~ block_literal? ~ modifier_suffix?
fn walk_command_call(mut items: Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    // The first item(s) before command_args form the callee postfix expression.
    let cmd_pos = items.iter().position(|p| p.as_rule() == Rule::command_args).unwrap();

    // Build the callee from the postfix pair(s) before command_args
    let callee_pairs: Vec<Pair<Rule>> = items.drain(..cmd_pos).collect();
    let callee = if callee_pairs.len() == 1 {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else if !callee_pairs.is_empty() {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else {
        return Err("Command call missing callee".into());
    };

    // Now items[0] = command_args (same structure as call_args: contains call_arg children)
    let cmd_args_pair = items.remove(0);
    let mut args = walk_call_args(cmd_args_pair)?;

    // Optional block literal
    if !items.is_empty() && items[0].as_rule() == Rule::block_literal {
        let blk = items.remove(0);
        let lambda = walk_block_literal(blk)?;
        args.push(Argument::positional(lambda));
    }

    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    });

    let stmt = StmtKind::Expr(call_expr);
    maybe_wrap_modifier(stmt, &mut items)
}

/// Wrap a statement in an if/unless/while/until modifier if present
fn maybe_wrap_modifier(stmt: StmtKind, rest: &mut Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    let mod_pos = rest.iter().position(|p| p.as_rule() == Rule::modifier_suffix);
    let mod_pair = match mod_pos {
        Some(pos) => rest.remove(pos),
        None => return Ok(stmt),
    };
    let mut mod_inner = mod_pair.into_inner();
    let kw = match mod_inner.next() {
        Some(k) => k,
        None => return Ok(stmt),
    };
    let cond_pair = mod_inner.next()
        .ok_or_else(|| "modifier_suffix missing condition".to_string())?;
    let cond = walk_expression(cond_pair)?;
    let body_stmt = Statement::new(stmt);
    match kw.as_rule() {
        Rule::if_kw => Ok(StmtKind::If {
            cond,
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::unless_kw => Ok(StmtKind::If {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::while_kw => Ok(StmtKind::While {
            cond,
            body: vec![body_stmt],
            else_body: None,
        }),
        Rule::until_kw => Ok(StmtKind::While {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            body: vec![body_stmt],
            else_body: None,
        }),
        _ => Ok(StmtKind::Expr(Expression::null())),
    }
}

// ── Require (import) ────────────────────────────────────────────────────────

fn walk_require(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let text = pair.as_str();
    let _is_relative = text.starts_with("require_relative");

    let mut path = String::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr_text = p.as_str().trim();
            // Strip quotes
            path = if expr_text.starts_with('"') || expr_text.starts_with('\'') {
                expr_text[1..expr_text.len()-1].to_string()
            } else {
                expr_text.to_string()
            };
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias: None },
        span,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Body (list of statements)
// ════════════════════════════════════════════════════════════════════════════

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::NEWLINE => {}
            _ => {
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    stmts.push(stmt);
                }
            }
        }
    }
    Ok(stmts)
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
        Rule::integer_literal => parse_ruby_int(pair.as_str()),
        Rule::float_literal => parse_ruby_float(pair.as_str()),
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(parse_ruby_string(pair.as_str())))),
        Rule::interpolated_string => walk_interpolated_string(pair),
        Rule::heredoc => Ok(ExprKind::Lit(Literal::Str(parse_heredoc(pair.as_str())))),
        Rule::symbol => Ok(ExprKind::Lit(Literal::Str(pair.as_str()[1..].to_string()))),
        Rule::regex_literal => Ok(ExprKind::Lit(Literal::Str(pair.as_str().to_string()))),
        Rule::percent_literal => Ok(walk_percent_literal(pair.as_str())),

        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::nil_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::self_kw => Ok(ExprKind::This),

        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant_path => Ok(ExprKind::Ident(pair.as_str().to_string())),

        // Instance var @x → self._rb_x  (prefixed to avoid collision with method bindings)
        Rule::instance_var => {
            let name = &pair.as_str()[1..]; // strip @
            Ok(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: format!("_rb_{}", name),
                null_safe: false,
            })
        }
        // Class var @@x → ident (treated as class-level variable)
        Rule::class_var => {
            let name = &pair.as_str()[2..]; // strip @@
            Ok(ExprKind::Ident(format!("_cls_{}", name)))
        }
        // Global var $x → ident
        Rule::global_var => {
            let name = &pair.as_str()[1..]; // strip $
            Ok(ExprKind::Ident(format!("_global_{}", name)))
        }

        // ── Expression wrappers ─────────────────────────────────────────
        Rule::expression => walk_expression_inner(pair),
        Rule::ternary_expr => walk_ternary(pair),
        Rule::or_expr | Rule::and_expr | Rule::not_expr |
        Rule::comparison | Rule::bitor_expr | Rule::bitxor_expr |
        Rule::bitand_expr | Rule::shift_expr | Rule::range_expr |
        Rule::additive | Rule::multiplicative | Rule::unary => walk_infix_or_unwrap(pair),

        Rule::postfix => walk_postfix(pair),
        Rule::primary => walk_primary(pair),
        Rule::expression_list => walk_expr_list_kind(pair),

        // ── Special expressions ─────────────────────────────────────────
        Rule::yield_expr => walk_yield(pair),
        Rule::defined_expr => walk_defined(pair),
        Rule::super_expr => walk_super(pair),
        Rule::block_given_expr => Ok(ExprKind::Lit(Literal::Bool(true))), // simplification
        Rule::lambda_literal => walk_lambda(pair),
        Rule::proc_literal => walk_proc(pair),

        // ── If/Unless/Begin as expression ───────────────────────────────
        Rule::if_expr => walk_if_expr(pair),
        Rule::unless_expr => walk_unless_expr(pair),
        Rule::begin_expr => walk_begin_expr(pair),

        Rule::array_inner => walk_array_inner(pair),
        Rule::hash_inner => walk_hash_inner(pair),

        Rule::NEWLINE => Ok(ExprKind::Lit(Literal::Null)),

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Expression inner (handles inline_rescue) ────────────────────────────────

fn walk_expression_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // expression = ternary_expr ~ inline_rescue?
    let expr = walk_expression(inner.remove(0))?;
    // If there's an inline_rescue, wrap in try
    if !inner.is_empty() && inner[0].as_rule() == Rule::inline_rescue {
        let rescue_inner: Vec<Pair<Rule>> = inner.remove(0).into_inner().collect();
        let _rescue_val = if let Some(rp) = rescue_inner.into_iter().next() {
            walk_expression(rp)?
        } else {
            Expression::null()
        };
        // Emit: (begin expr rescue => rescue_val end) as a ternary
        // Simplification: just return the expr (rescue is error handling)
        return Ok(expr.kind);
    }
    Ok(expr.kind)
}

// ── Ternary ─────────────────────────────────────────────────────────────────

fn walk_ternary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // cond ? then : else
    if inner.len() >= 3 {
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

// ── Infix / precedence unwrap ───────────────────────────────────────────────

fn walk_infix_or_unwrap(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    match rule {
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::not_expr => {
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            Ok(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(operand) })
        }
        Rule::comparison => {
            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i < inner.len() {
                if inner[i].as_rule() == Rule::comparison_op {
                    let op = parse_comparison_op(inner[i].as_str().trim());
                    i += 1;
                    if i < inner.len() {
                        let right = walk_expression(inner[i].clone())?;
                        i += 1;
                        left = Expression::new(ExprKind::Binary {
                            op, left: Box::new(left), right: Box::new(right),
                        });
                    }
                } else {
                    i += 1;
                }
            }
            Ok(left.kind)
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::range_expr => walk_range(inner),
        Rule::additive => walk_binary_chain_with_ops(inner),
        Rule::multiplicative => walk_ruby_multiplicative(inner),
        Rule::unary => {
            let op_str = inner[0].as_str().trim();
            let operand = walk_expression(inner.pop().ok_or("Empty unary")?)?;
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary { op, expr: Box::new(operand) })
        }
        _ => {
            if !inner.is_empty() {
                walk_expr_kind(inner.remove(0))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }
    }
}

fn walk_binary_chain(mut items: Vec<Pair<Rule>>, op_fn: impl Fn(&str) -> BinOp) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_expression_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            let op = op_fn("");
            left = Expression::new(ExprKind::Binary {
                op, left: Box::new(left), right: Box::new(right),
            });
        }
    }
    Ok(left.kind)
}

/// Ruby `*` is dynamic (string repeat OR numeric mul), same as Python.
fn walk_ruby_multiplicative(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                let op = parse_binop(op_str);
                left = Expression::new(ExprKind::Binary {
                    op, left: Box::new(left), right: Box::new(right),
                });
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn walk_binary_chain_with_ops(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim());
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op, left: Box::new(left), right: Box::new(right),
                });
            }
        } else if is_expression_rule(p.as_rule()) {
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = Expression::new(ExprKind::Binary {
                op: BinOp::Add, left: Box::new(left), right: Box::new(right),
            });
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

// ── Range ───────────────────────────────────────────────────────────────────

fn walk_range(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    if items.len() == 1 {
        return walk_expr_kind(items.remove(0));
    }
    let start = walk_expression(items.remove(0))?;
    // Find range_op
    let mut inclusive = true;
    let mut end_idx = 0;
    for (i, p) in items.iter().enumerate() {
        if p.as_rule() == Rule::range_op {
            inclusive = p.as_str() == "..";
            end_idx = i + 1;
            break;
        }
    }
    if end_idx < items.len() {
        let end = walk_expression(items.remove(end_idx))?;
        if inclusive {
            // Compiler ignores inclusive flag — normalize to exclusive: end + 1
            let end_plus_one = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(end),
                right: Box::new(Expression::int(1)),
            });
            Ok(ExprKind::Range {
                start: Box::new(start),
                end: Box::new(end_plus_one),
                inclusive: false,
            })
        } else {
            Ok(ExprKind::Range {
                start: Box::new(start),
                end: Box::new(end),
                inclusive: false,
            })
        }
    } else {
        Ok(start.kind)
    }
}

// ── Postfix (call, member, subscript, block) ────────────────────────────────

fn walk_postfix(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty postfix")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() == Rule::postfix_chain {
            expr = walk_postfix_chain(expr, chain)?;
        }
    }
    Ok(expr.kind)
}

fn walk_postfix_chain(expr: Expression, chain: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = chain.into_inner().collect();

    if children.is_empty() {
        // bare () call
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(expr),
            args: Vec::new(),
            optional: false,
        }));
    }

    let first_rule = children[0].as_rule();

    match first_rule {
        Rule::method_name_id => {
            // Method call: .method or &.method
            let method_name = children[0].as_str().to_string();
            let null_safe = children.iter().any(|c| c.as_str() == "&.");

            // Check if there are call args
            let args = children.iter()
                .find(|c| c.as_rule() == Rule::call_args)
                .map(|c| walk_call_args(c.clone()))
                .transpose()?
                .unwrap_or_default();

            // Check for trailing block
            let block = children.iter()
                .find(|c| c.as_rule() == Rule::block_literal)
                .map(|c| walk_block_literal(c.clone()))
                .transpose()?;

            let mut final_args = args;
            if let Some(block_lambda) = block {
                final_args.push(Argument::positional(block_lambda));
            }

            // Normalize .new() → ExprKind::New (constructor call)
            if method_name == "new" {
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(expr),
                    args: final_args,
                }));
            }

            // Normalize .call() → direct call (lambda/proc invocation)
            if method_name == "call" {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args: final_args,
                    optional: false,
                }));
            }

            // Normalize .first → Index(expr, 0) — pure bytecode, no host call
            if method_name == "first" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(Expression::int(0)),
                    null_safe: false,
                }));
            }

            // Normalize .last → Index(expr, -1) — pure bytecode
            if method_name == "last" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(Expression::int(-1)),
                    null_safe: false,
                }));
            }

            let member = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe,
            });

            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args: final_args,
                optional: false,
            }))
        }
        Rule::constant => {
            // Scope resolution: ::Constant
            let const_name = children[0].as_str();
            Ok(Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: const_name.to_string(),
                null_safe: false,
            }))
        }
        Rule::call_args => {
            // Bare call: expr(args)
            let args = walk_call_args(children.into_iter().next().unwrap())?;
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args,
                optional: false,
            }))
        }
        Rule::expression_list => {
            // Subscript: expr[index]
            let index = walk_expr_list_single(children.into_iter().next().unwrap())?;
            Ok(Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(index),
                null_safe: false,
            }))
        }
        Rule::block_literal => {
            // Trailing block on its own (e.g., `array.each { |x| ... }`)
            // The method call should already be formed; this adds the block as arg
            if let ExprKind::Call { callee, mut args, optional } = expr.kind {
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(block_lambda));
                Ok(Expression::new(ExprKind::Call { callee, args, optional }))
            } else {
                // Bare block on expression — treat as call with block
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args: vec![Argument::positional(block_lambda)],
                    optional: false,
                }))
            }
        }
        _ => {
            // Try to interpret as subscript or call
            if is_expression_rule(first_rule) {
                let index = walk_expression(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
                    null_safe: false,
                }))
            } else {
                Ok(expr)
            }
        }
    }
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.is_empty() { continue; }

            let first_text = children[0].as_str();

            if first_text == "**" {
                // Double splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument { value: val, name: None, by_ref: false, spread: true });
                }
            } else if first_text == "*" {
                // Splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument { value: val, name: None, by_ref: false, spread: true });
                }
            } else if first_text == "&" {
                // Block arg
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument::positional(val));
                }
            } else if children.len() >= 2 && children[0].as_rule() == Rule::identifier {
                // Check if keyword arg: identifier ":" expression
                let has_colon = children.iter().any(|c| c.as_str() == ":");
                if has_colon {
                    let name = children[0].as_str().to_string();
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    args.push(Argument { value: val, name: Some(name), by_ref: false, spread: false });
                } else {
                    let val = walk_expression(children.into_iter().next().unwrap())?;
                    args.push(Argument::positional(val));
                }
            } else {
                let val = walk_expression(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(val));
            }
        }
    }
    Ok(args)
}

// ── Block literal ───────────────────────────────────────────────────────────

fn walk_block_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block | Rule::brace_block => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::block_params => {
                            params = walk_block_params(bp)?;
                        }
                        Rule::body => {
                            body = walk_body(bp)?;
                        }
                        _ => {
                            // Statements directly in brace_block
                            let stmt = walk_statement(bp)?;
                            if !matches!(stmt.kind, StmtKind::Empty) {
                                body.push(stmt);
                            }
                        }
                    }
                }
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

/// Ruby implicit return: last expression in a body becomes a Return.
fn apply_implicit_return(body: &mut Vec<Statement>) {
    if let Some(last) = body.last_mut() {
        if matches!(&last.kind, StmtKind::Expr(_)) {
            if let StmtKind::Expr(e) = std::mem::replace(&mut last.kind, StmtKind::Empty) {
                last.kind = StmtKind::Return(Some(e));
            }
        }
    }
}

fn walk_block_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_param_list {
            for bp in p.into_inner() {
                if bp.as_rule() == Rule::block_param_item {
                    let inner = bp.into_inner().next();
                    if let Some(item) = inner {
                        match item.as_rule() {
                            Rule::splat_param => {
                                let name = item.into_inner()
                                    .find(|c| c.as_rule() == Rule::identifier)
                                    .map(|c| c.as_str().to_string())
                                    .unwrap_or_default();
                                params.push(Param {
                                    name,
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: true,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                            Rule::identifier => {
                                params.push(Param {
                                    name: item.as_str().to_string(),
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
            }
        }
    }
    Ok(params)
}

// ── Primary ─────────────────────────────────────────────────────────────────

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    if inner.is_empty() {
        return Ok(ExprKind::Lit(Literal::Null));
    }

    let first = &inner[0];
    match first.as_rule() {
        Rule::array_inner => {
            // Array literal [...]
            walk_array_inner(inner.remove(0))
        }
        Rule::hash_inner => {
            // Hash literal {...}
            walk_hash_inner(inner.remove(0))
        }
        _ => walk_expr_kind(inner.remove(0)),
    }
}

fn walk_array_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements = pair.into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            Ok(ArrayElement { key: None, value: val, spread: false, by_ref: false })
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_hash_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut props = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::hash_pair {
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.len() >= 2 {
                // Could be hash rocket (key => val) or symbol shorthand (key: val)
                let first = &children[0];
                if first.as_rule() == Rule::identifier && children.len() == 2 {
                    // Symbol shorthand: key: val
                    let key = Expression::new(ExprKind::Lit(Literal::Str(first.as_str().to_string())));
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                } else {
                    let key = walk_expression(children[0].clone())?;
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                }
            } else if children.len() == 1 {
                // **expr (double splat)
                let val = walk_expression(children.into_iter().next().unwrap())?;
                props.push(ObjectProperty::Spread(val));
            }
        }
    }
    Ok(ExprKind::Object(props))
}

// ── Interpolated string ─────────────────────────────────────────────────────

fn walk_interpolated_string(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::interp_start | Rule::interp_end => {}
            Rule::interp_text => {
                let text = p.as_str()
                    .replace("\\n", "\n")
                    .replace("\\t", "\t")
                    .replace("\\r", "\r")
                    .replace("\\\\", "\\")
                    .replace("\\\"", "\"");
                parts.push(InterpolPart::Text(text));
            }
            Rule::interp_escape => {
                let s = p.as_str();
                let ch = if s.len() >= 2 {
                    match s.chars().nth(1) {
                        Some('n') => "\n",
                        Some('t') => "\t",
                        Some('r') => "\r",
                        Some('\\') => "\\",
                        Some('"') => "\"",
                        Some('#') => "#",
                        _ => s,
                    }
                } else { s };
                parts.push(InterpolPart::Text(ch.to_string()));
            }
            Rule::interp_expr => {
                for ip in p.into_inner() {
                    if is_expression_rule(ip.as_rule()) {
                        parts.push(InterpolPart::Expr(walk_expression(ip)?));
                    }
                }
            }
            _ => {}
        }
    }

    // Optimize: if only text parts, concat into single string
    if parts.iter().all(|p| matches!(p, InterpolPart::Text(_))) {
        let s: String = parts.iter().map(|p| match p {
            InterpolPart::Text(t) => t.as_str(),
            _ => "",
        }).collect();
        return Ok(ExprKind::Lit(Literal::Str(s)));
    }

    Ok(ExprKind::Interpolation(parts))
}

// ── Lambda ──────────────────────────────────────────────────────────────────

fn walk_lambda(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_param_list(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {
                // Statements in lambda brace body
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    body.push(stmt);
                }
            }
        }
    }

    apply_implicit_return(&mut body);

    Ok(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

fn walk_proc(pair: Pair<Rule>) -> Result<ExprKind, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_literal {
            let lambda = walk_block_literal(p)?;
            return Ok(lambda.kind);
        }
    }
    Ok(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(Vec::new()),
        is_async: false,
        captures: Vec::new(),
    })
}

// ── Yield ───────────────────────────────────────────────────────────────────

fn walk_yield(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    args.push(walk_expression(ep)?);
                }
            }
        } else if is_expression_rule(p.as_rule()) {
            args.push(walk_expression(p)?);
        }
    }
    // Ruby yield calls the block; emit as Yield for now
    if args.is_empty() {
        Ok(ExprKind::Yield(None))
    } else if args.len() == 1 {
        Ok(ExprKind::Yield(Some(Box::new(args.into_iter().next().unwrap()))))
    } else {
        Ok(ExprKind::Yield(Some(Box::new(Expression::new(ExprKind::Array(
            args.into_iter().map(|a| ArrayElement { key: None, value: a, spread: false, by_ref: false }).collect()
        ))))))
    }
}

// ── Defined? ────────────────────────────────────────────────────────────────

fn walk_defined(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // defined?(expr) → check if expr is defined, simplify to !nil
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr = walk_expression(p)?;
            return Ok(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(expr),
                right: Box::new(Expression::null()),
            });
        }
    }
    Ok(ExprKind::Lit(Literal::Bool(false)))
}

// ── Super ───────────────────────────────────────────────────────────────────

fn walk_super(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_args {
            args = walk_call_args(p)?;
        }
    }
    Ok(ExprKind::SuperCall { method: None, args })
}

// ── If/Unless as expression ─────────────────────────────────────────────────

fn walk_if_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_if(pair)?;
    // Wrap as a ternary-like expression
    if let StmtKind::If { cond, then_body, else_body, .. } = kind {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_unless_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_unless(pair)?;
    if let StmtKind::If { cond, then_body, else_body, .. } = kind {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_begin_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // begin..rescue..end as expression — just walk the body
    let kind = walk_begin(pair)?;
    if let StmtKind::Try { body, .. } = kind {
        Ok(body_to_expr(body).kind)
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

/// Convert a body (list of stmts) to a single expression (last statement value).
fn body_to_expr(mut stmts: Vec<Statement>) -> Expression {
    if stmts.is_empty() {
        return Expression::null();
    }
    let last = stmts.pop().unwrap();
    match last.kind {
        StmtKind::Expr(e) => e,
        StmtKind::Return(Some(e)) => e,
        _ => Expression::null(),
    }
}

// ── Expression list ─────────────────────────────────────────────────────────

fn walk_expr_list_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expr_kind(inner.into_iter().next().unwrap())
    } else if inner.is_empty() {
        Ok(ExprKind::Lit(Literal::Null))
    } else {
        let exprs = inner.into_iter().map(walk_expression).collect::<Result<Vec<_>, _>>()?;
        Ok(ExprKind::Array(
            exprs.into_iter().map(|e| ArrayElement { key: None, value: e, spread: false, by_ref: false }).collect()
        ))
    }
}

fn walk_expr_list_single(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expression(inner.remove(0))
    } else if inner.is_empty() {
        Ok(Expression::null())
    } else {
        let exprs = inner.into_iter().map(walk_expression).collect::<Result<Vec<_>, _>>()?;
        Ok(Expression::new(ExprKind::Array(
            exprs.into_iter().map(|e| ArrayElement { key: None, value: e, spread: false, by_ref: false }).collect()
        )))
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32,
    }
}

fn negate(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn next_meaningful<'a>(iter: &mut impl Iterator<Item = Pair<'a, Rule>>) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        match p.as_rule() {
            Rule::NEWLINE | Rule::then_kw | Rule::do_kw | Rule::in_kw => continue,
            _ => return Ok(p),
        }
    }
    Err("No more meaningful pairs".into())
}

fn next_rule<'a>(iter: &mut impl Iterator<Item = Pair<'a, Rule>>, rule: Rule) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule<'a>(iter: impl Iterator<Item = Pair<'a, Rule>>, rule: Rule) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule_from_iter<'a>(iter: &mut impl Iterator<Item = Pair<'a, Rule>>, rule: Rule) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn is_expression_rule(rule: Rule) -> bool {
    matches!(rule,
        Rule::expression | Rule::expression_list |
        Rule::ternary_expr | Rule::or_expr | Rule::and_expr | Rule::not_expr |
        Rule::comparison | Rule::bitor_expr | Rule::bitxor_expr |
        Rule::bitand_expr | Rule::shift_expr | Rule::range_expr |
        Rule::additive | Rule::multiplicative | Rule::unary |
        Rule::postfix | Rule::primary |
        Rule::integer_literal | Rule::float_literal |
        Rule::string_literal | Rule::interpolated_string | Rule::heredoc |
        Rule::symbol | Rule::regex_literal | Rule::percent_literal |
        Rule::true_kw | Rule::false_kw | Rule::nil_kw | Rule::self_kw |
        Rule::identifier | Rule::constant | Rule::constant_path |
        Rule::instance_var | Rule::class_var | Rule::global_var |
        Rule::yield_expr | Rule::defined_expr | Rule::super_expr |
        Rule::block_given_expr | Rule::lambda_literal | Rule::proc_literal |
        Rule::if_expr | Rule::unless_expr | Rule::begin_expr
    )
}

fn is_op_rule(rule: Rule) -> bool {
    matches!(rule,
        Rule::additive_op | Rule::multiplicative_op | Rule::shift_op |
        Rule::comparison_op | Rule::range_op | Rule::aug_assign_op
    )
}

fn parse_comparison_op(s: &str) -> BinOp {
    match s {
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "<=>" => BinOp::Spaceship,
        "===" => BinOp::StrictEq,
        _ => BinOp::Eq,
    }
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn parse_ruby_int(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    if s.starts_with("0x") || s.starts_with("0X") {
        Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 16).unwrap_or(0))))
    } else if s.starts_with("0o") || s.starts_with("0O") {
        Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 8).unwrap_or(0))))
    } else if s.starts_with("0b") || s.starts_with("0B") {
        Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 2).unwrap_or(0))))
    } else {
        Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
    }
}

fn parse_ruby_float(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    Ok(ExprKind::Lit(Literal::Float(s.parse().unwrap_or(0.0))))
}

fn parse_ruby_string(s: &str) -> String {
    let s = if s.starts_with("'''") {
        &s[3..s.len()-3]
    } else if s.starts_with('\'') {
        &s[1..s.len()-1]
    } else {
        s
    };
    // Single-quoted strings: only \\ and \' are escapes
    s.replace("\\'", "'")
     .replace("\\\\", "\\")
}

fn parse_heredoc(s: &str) -> String {
    // <<~TAG\ncontent\nTAG  or  <<TAG\ncontent\nTAG
    let squiggly = s.starts_with("<<~");
    let prefix_len = if squiggly { 3 } else { 2 };
    let rest = &s[prefix_len..];
    // Find the tag name (up to newline)
    if let Some(nl) = rest.find('\n') {
        let tag = rest[..nl].trim();
        let content = &rest[nl+1..];
        // Strip trailing TAG line
        let body = if let Some(pos) = content.rfind(tag) {
            &content[..pos]
        } else {
            content
        };
        if squiggly {
            // Strip common leading whitespace
            let lines: Vec<&str> = body.lines().collect();
            let min_indent = lines.iter()
                .filter(|l| !l.trim().is_empty())
                .map(|l| l.len() - l.trim_start().len())
                .min()
                .unwrap_or(0);
            lines.iter()
                .map(|l| if l.len() > min_indent { &l[min_indent..] } else { l.trim() })
                .collect::<Vec<_>>()
                .join("\n")
        } else {
            body.to_string()
        }
    } else {
        s.to_string()
    }
}

fn walk_percent_literal(s: &str) -> ExprKind {
    // %w[a b c] → array of strings
    // %i[a b c] → array of symbols (strings)
    // %q[...] → single-quoted string
    // %Q[...] or %[...] → double-quoted string
    let (kind, rest) = if s.starts_with("%w") || s.starts_with("%i") {
        ("array", &s[2..])
    } else if s.starts_with("%q") || s.starts_with("%Q") {
        ("string", &s[2..])
    } else {
        ("string", &s[1..])
    };

    // Strip delimiters
    let body = if rest.starts_with('[') {
        &rest[1..rest.len()-1]
    } else if rest.starts_with('(') {
        &rest[1..rest.len()-1]
    } else if rest.starts_with('{') {
        &rest[1..rest.len()-1]
    } else if rest.starts_with('<') {
        &rest[1..rest.len()-1]
    } else {
        rest
    };

    if kind == "array" {
        let words: Vec<ArrayElement> = body.split_whitespace()
            .map(|w| ArrayElement {
                key: None,
                value: Expression::new(ExprKind::Lit(Literal::Str(w.to_string()))),
                spread: false,
                by_ref: false,
            })
            .collect();
        ExprKind::Array(words)
    } else {
        ExprKind::Lit(Literal::Str(body.to_string()))
    }
}
