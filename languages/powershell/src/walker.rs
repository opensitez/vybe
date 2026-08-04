//! PowerShell walker.
//!
//! Transforms Pest parse trees into the shared JS-shaped AST used by compiler
//! primitives.

use pest::iterators::Pair;
use pest::Parser;
use vybe_ast::*;
use super::Rule;

use std::collections::VecDeque;

pub fn parse(source: &str) -> Result<Module, String> {
    let source = normalize_source(source);
    let mut pairs =
        super::PowerShellParser::parse(Rule::program, &source).map_err(|e| format!("Parse error: {e}"))?;

    let pair = pairs
        .next()
        .ok_or_else(|| "No PowerShell parse root".to_string())?;

    let mut body = Vec::new();

    for child in pair.into_inner() {
        if let Some(stmt) = parse_statement(child)? {
            body.push(stmt);
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
    })
}

fn normalize_source(source: &str) -> String {
    let src = source.trim_start_matches('\u{feff}').replace("\r\n", "\n");
    let mut out = String::new();
    let mut lines = src.lines().peekable();

    while let Some(raw) = lines.next() {
        let trimmed = raw.trim_end();
        if trimmed.ends_with('`') {
            out.push_str(trimmed.trim_end_matches('`'));
            if lines.peek().is_some() {
                out.push(' ');
            }
        } else {
            out.push_str(trimmed);
            if lines.peek().is_some() {
                out.push('\n');
            }
        }
    }

    out
}

fn parse_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let mut queue = VecDeque::new();
    queue.push_back(pair);

    while let Some(pair) = queue.pop_front() {
        match pair.as_rule() {
            Rule::statement => {
                let text = pair.as_str().trim();
                if looks_like_command_invocation(text) && text.contains(' ') {
                    if let Some(expr) = parse_command_line(text) {
                        return Ok(Some(Statement::new(StmtKind::Expr(expr))));
                    }
                }

                let children: Vec<Pair<Rule>> = pair.into_inner().collect();
                let mut tokens = Vec::new();
                for child in &children {
                    collect_command_tokens(child, &mut tokens);
                }

                if let Some(expr) = parse_command_tokens_as_expr(&tokens) {
                    return Ok(Some(Statement::new(StmtKind::Expr(expr))));
                }

                for child in children {
                    queue.push_back(child);
                }
            }
            Rule::COMMENT => return Ok(None),
            Rule::namespace_decl => return Ok(Some(parse_namespace_decl(pair)?)),
            Rule::class_decl => return Ok(Some(parse_class_decl(pair)?)),
            Rule::function_decl => return Ok(Some(parse_function_decl(pair)?)),
            Rule::if_stmt => return Ok(Some(parse_if_stmt(pair)?)),
            Rule::switch_stmt => return Ok(Some(parse_switch_stmt(pair)?)),
            Rule::foreach_stmt => return Ok(Some(parse_foreach_stmt(pair)?)),
            Rule::for_stmt => return Ok(Some(parse_for_stmt(pair)?)),
            Rule::while_stmt => return Ok(Some(parse_while_stmt(pair)?)),
            Rule::do_while_stmt => return Ok(Some(parse_do_while_stmt(pair)?)),
            Rule::try_stmt => return Ok(Some(parse_try_stmt(pair)?)),
            Rule::return_stmt => return Ok(Some(parse_return_stmt(pair))),
            Rule::throw_stmt => return Ok(Some(parse_throw_stmt(pair))),
            Rule::break_stmt => return Ok(Some(parse_break_stmt(pair))),
            Rule::continue_stmt => return Ok(Some(parse_continue_stmt(pair))),
            Rule::param_stmt => return Ok(None),
            Rule::using_stmt => return Ok(None),
            Rule::assignment_stmt => return Ok(Some(parse_assignment_statement(pair))),
            Rule::increment_stmt => return Ok(Some(parse_increment_statement(pair))),
            Rule::command_stmt => return Ok(Some(parse_command_statement(pair))),
            _ => {
                if let Some(stmt) = parse_statement_fallback(pair)? {
                    return Ok(Some(stmt));
                }
            }
        }
    }

    Ok(None)
}

fn parse_statement_fallback(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let body = pair.as_str().trim();
    if body.is_empty() {
        return Ok(None);
    }
    if pair.as_rule() == Rule::expr_text {
        return Ok(Some(Statement::new(StmtKind::Expr(expr_from_text(body)))));
    }
    Ok(None)
}

fn parse_namespace_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::namespace_name | Rule::DOTTED_NAME => {
                name = child.as_str().trim().to_string();
            }
            Rule::block => body = parse_block_statements(child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::NamespaceDecl { name, body }))
}

fn parse_block_statements(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for child in pair.into_inner() {
        if let Some(stmt) = parse_statement(child)? {
            body.push(stmt);
        }
    }
    Ok(body)
}

fn parse_class_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::IDENT => {
                name = child.as_str().trim().to_string();
            }
            Rule::class_heritage => {
                for chunk in child
                    .as_str()
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                {
                    parents.push(chunk);
                }
            }
            Rule::class_body => {
                members = parse_class_body(child)?;
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers::default(),
        decorators: Vec::new(),
    }))
}

fn parse_class_body(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::class_member_decl && let Some(member) = parse_class_member(child)? {
            members.push(member);
        }
    }
    Ok(members)
}

fn parse_class_member(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    let mut inner = pair.into_inner();
    if let Some(child) = inner.next() {
        return match child.as_rule() {
            Rule::class_function_decl => {
                let statement = parse_function_decl(child)?;
                Ok(Some(ClassMember::Method(Box::new(statement))))
            }
            Rule::class_field_decl => parse_class_field_decl(child).map(Some),
            Rule::constructor_decl => parse_constructor_decl(child).map(Some),
            _ => Ok(None),
        };
    }
    Ok(None)
}

fn parse_class_field_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut init = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::var_ref => {
                name = child.as_str().trim_start_matches('$').to_string();
            }
            Rule::expr_text => {
                init = Some(expr_from_text(child.as_str()));
            }
            _ => {}
        }
    }

    Ok(ClassMember::Field {
        name,
        type_hint: None,
        init,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
    })
}

fn parse_constructor_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::function_params => params = parse_function_params(child),
            Rule::block => body = parse_block_with_function_params(child, &mut params)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

fn parse_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::function_name => {
                name = child.as_str().trim().to_string();
            }
            Rule::function_params => params = parse_function_params(child),
            Rule::block => body = parse_block_with_function_params(child, &mut params)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn parse_function_params(pair: Pair<Rule>) -> Vec<Param> {
    let mut out = Vec::new();
    let mut param_nodes = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::function_param_list || child.as_rule() == Rule::function_param {
            param_nodes.push(child);
        }
    }

    if param_nodes.is_empty() {
        return out;
    }

    for raw in param_nodes.into_iter().flat_map(|node| node.into_inner()) {
        if raw.as_rule() != Rule::function_param {
            continue;
        }

        let mut name = String::new();
        let mut default = None;
        let mut type_hint = None;

        for piece in raw.into_inner() {
            match piece.as_rule() {
                Rule::type_hint => {
                    let inner = piece.as_str().trim();
                    let inner = inner.trim_start_matches('[').trim_end_matches(']');
                    let inner = inner.trim();
                    if !inner.is_empty() {
                        type_hint = Some(inner.to_string());
                    }
                }
                Rule::var_ref | Rule::IDENT => {
                    name = piece.as_str().trim_start_matches('$').to_string();
                }
                Rule::expr_text => {
                    default = Some(expr_from_text(piece.as_str()));
                }
                _ => {}
            }
        }

        if !name.is_empty() {
            // Keep `default` readable for metadata checks before moving into `Param`.
            let is_optional = default.as_ref().is_some();
            out.push(Param {
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
    out
}

fn parse_param_stmt(pair: Pair<Rule>, out: &mut Vec<Param>) {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::function_param_list {
            out.extend(parse_function_params(child));
        }
    }
}

fn parse_block_with_function_params(
    pair: Pair<Rule>,
    params: &mut Vec<Param>,
) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    let mut consume_param_stmt = true;

    for child in pair.into_inner() {
        if child.as_rule() != Rule::statement {
            continue;
        }

        if consume_param_stmt {
            let mut consumed = false;
            for inner in child.clone().into_inner() {
                if inner.as_rule() == Rule::param_stmt {
                    parse_param_stmt(inner, params);
                    consumed = true;
                    break;
                }
            }
            if consumed {
                continue;
            }
            consume_param_stmt = false;
        }

        if let Some(stmt) = parse_statement(child)? {
            body.push(stmt);
        }
    }

    Ok(body)
}

fn parse_if_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::condition_expr => cond = Some(parse_condition(child.as_str())),
            Rule::expr_text => cond = Some(expr_from_text(child.as_str())),
            Rule::block => then_body = parse_block_statements(child)?,
            Rule::elseif_stmt => {
                let mut branch_cond = None;
                let mut branch_body = Vec::new();
                for part in child.into_inner() {
                    if part.as_rule() == Rule::condition_expr {
                        branch_cond = Some(parse_condition(part.as_str()));
                    } else if part.as_rule() == Rule::expr_text {
                        branch_cond = Some(expr_from_text(part.as_str()));
                    }
                    if part.as_rule() == Rule::block {
                        branch_body = parse_block_statements(part)?;
                    }
                }
                if let Some(branch_cond) = branch_cond {
                    elifs.push((branch_cond, branch_body));
                }
            }
            Rule::else_stmt => {
                if let Some(block) = child.into_inner().find(|c| c.as_rule() == Rule::block) {
                    else_body = Some(parse_block_statements(block)?);
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(Expression::null),
        then_body,
        elifs,
        else_body,
    }))
}

fn parse_condition(raw: &str) -> Expression {
    let text = raw.trim();
    if is_fully_wrapped(text, '(', ')') {
        return expr_from_text(strip_outer_parentheses(text));
    }

    expr_from_text(text)
}

fn parse_foreach_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut iter_var = String::new();
    let mut iter = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::var_ref => iter_var = child.as_str().trim_start_matches('$').to_string(),
            Rule::expr_text => iter = Some(expr_from_text(child.as_str())),
            Rule::block => body = parse_block_statements(child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::ForIn {
        var: iter_var,
        key: None,
        iter: iter.unwrap_or_else(Expression::null),
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
}

fn parse_switch_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut expr = None;
    let mut cases = Vec::new();
    let mut default = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::expr_text => {
                expr = Some(expr_from_text(child.as_str()));
            }
            Rule::condition_expr => {
                expr = Some(expr_from_text(child.as_str()));
            }
            Rule::switch_case => {
                let mut branch_body = None;
                let mut branch_conditions = Vec::new();
                let is_default = child
                    .as_str()
                    .trim_start()
                    .starts_with("default");

                for part in child.into_inner() {
                    match part.as_rule() {
                        Rule::expr_text if !is_default => {
                            branch_conditions.push(CaseCondition::Value(expr_from_text(part.as_str())));
                        }
                        Rule::switch_case_body => {
                            branch_body = Some(parse_block_statements(part)?);
                        }
                        Rule::switch_default_case => {
                            for inner in part.into_inner() {
                                if inner.as_rule() == Rule::switch_case_body {
                                    branch_body = Some(parse_block_statements(inner)?);
                                    break;
                                }
                            }
                        }
                        Rule::switch_case_value => {
                            for inner in part.into_inner() {
                                match inner.as_rule() {
                                    Rule::expr_text => {
                                        branch_conditions.push(CaseCondition::Value(expr_from_text(inner.as_str())));
                                    }
                                    Rule::switch_case_body => {
                                        branch_body = Some(parse_block_statements(inner)?);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }

                if is_default {
                    default = branch_body.or(Some(Vec::new()));
                    continue;
                }

                if let Some(body) = branch_body {
                    cases.push(SwitchCase {
                        conditions: branch_conditions,
                        body,
                    });
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::Switch {
        expr: expr.unwrap_or_else(Expression::null),
        cases,
        default,
    }))
}

fn parse_for_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut header = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::for_header => header = Some(child.as_str().to_string()),
            Rule::block => body = parse_block_statements(child)?,
            _ => {}
        }
    }

    let header = header.unwrap_or_default();
    let mut it = header.split(';').map(str::trim).collect::<Vec<_>>();
    if it.len() < 3 {
        it.resize(3, "");
    }

    let init = parse_for_part_expr(it[0]).filter(|expr| {
        !matches!(expr.kind, ExprKind::Ident(_)) || !expr.as_ident_empty()
    });
    let cond = parse_for_part_expr(it[1]).filter(|expr| {
        !matches!(expr.kind, ExprKind::Ident(_)) || !expr.as_ident_empty()
    });
    let update = parse_for_part_expr(it[2]).filter(|expr| {
        !matches!(expr.kind, ExprKind::Ident(_)) || !expr.as_ident_empty()
    });

    Ok(Statement::new(StmtKind::For {
        init: init.map(|expr| Box::new(Statement::new(StmtKind::Expr(expr)))),
        cond: cond,
        update: update,
        body,
    }))
}

fn parse_for_part_expr(text: &str) -> Option<Expression> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    if let Some((target, op, rhs)) = detect_assignment_like(text) {
        let value = expr_from_text(rhs.as_str());
        let lhs = assignment_target(&target);
        return Some(match op.as_str() {
            "=" => Expression::new(ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(value),
            }),
            "+=" => binary_assign(lhs, BinOp::Add, value),
            "-=" => binary_assign(lhs, BinOp::Sub, value),
            "*=" => binary_assign(lhs, BinOp::Mul, value),
            "/=" => binary_assign(lhs, BinOp::Div, value),
            "%=" => binary_assign(lhs, BinOp::Mod, value),
            "**=" => binary_assign(lhs, BinOp::Pow, value),
            "&=" => binary_assign(lhs, BinOp::BitAnd, value),
            "|=" => binary_assign(lhs, BinOp::BitOr, value),
            "^=" => binary_assign(lhs, BinOp::BitXor, value),
            "<<=" => binary_assign(lhs, BinOp::Shl, value),
            ">>=" => binary_assign(lhs, BinOp::Shr, value),
            "&&=" => binary_assign(lhs, BinOp::And, value),
            "||=" => binary_assign(lhs, BinOp::Or, value),
            "??=" => binary_assign(lhs, BinOp::NullCoalesce, value),
            _ => Expression::new(ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(value),
            }),
        });
    }
    if let Some(target_text) = text.strip_suffix("++") {
        let target = assignment_target(target_text.trim());
        Some(Expression::new(ExprKind::Assign {
            target: Box::new(target.clone()),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(target),
                right: Box::new(Expression::int(1)),
            })),
        }))
    } else if let Some(target_text) = text.strip_suffix("--") {
        let target = assignment_target(target_text.trim());
        Some(Expression::new(ExprKind::Assign {
            target: Box::new(target.clone()),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(target),
                right: Box::new(Expression::int(1)),
            })),
        }))
    } else {
        Some(expr_from_text(text))
    }
}

fn binary_assign(target: Expression, op: BinOp, value: Expression) -> Expression {
    Expression::new(ExprKind::Assign {
        target: Box::new(target.clone()),
        value: Box::new(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(target),
            right: Box::new(value),
        })),
    })
}

fn parse_while_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut cond = None;
    let mut body = Vec::new();

        for child in pair.into_inner() {
        if child.as_rule() == Rule::expr_text {
            cond = Some(expr_from_text(child.as_str()));
        }
        if child.as_rule() == Rule::condition_expr {
            cond = Some(parse_condition(child.as_str()));
        }
        if child.as_rule() == Rule::block {
            body = parse_block_statements(child)?;
        }
    }

    Ok(Statement::new(StmtKind::While {
        cond: cond.unwrap_or_else(Expression::null),
        body,
        else_body: None,
    }))
}

fn parse_do_while_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    let mut cond = None;
    let mut until = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::block => body = parse_block_statements(child)?,
            Rule::do_while_tail => {
                let text = child.as_str().trim().to_lowercase();
                until = text.starts_with("until");
                    for inner in child.into_inner() {
                    if inner.as_rule() == Rule::expr_text {
                        cond = Some(expr_from_text(inner.as_str()));
                    }
                    if inner.as_rule() == Rule::condition_expr {
                        cond = Some(parse_condition(inner.as_str()));
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::DoWhile {
        body,
        cond: cond.unwrap_or_else(Expression::null),
        until,
    }))
}

fn parse_try_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::block => body = parse_block_statements(child)?,
            Rule::catch_clause => catches.push(parse_catch_clause(child)?),
            Rule::finally_clause => finally = Some(parse_block_statements(child)?),
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    }))
}

fn parse_catch_clause(pair: Pair<Rule>) -> Result<CatchClause, String> {
    let mut var_name = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        if child.as_rule() == Rule::var_ref {
            let raw = child.as_str().trim().trim_start_matches('$');
            if !raw.is_empty() {
                var_name = Some(raw.to_string());
            }
        }
        if child.as_rule() == Rule::block {
            body = parse_block_statements(child)?;
        }
    }

    Ok(CatchClause {
        types: Vec::new(),
        var_name,
        stack_var: None,
        body,
        when_clause: None,
    })
}

fn parse_return_stmt(pair: Pair<Rule>) -> Statement {
    let expr = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::expr_text)
        .map(|c| expr_from_text(c.as_str()));
    Statement::new(StmtKind::Return(expr))
}

fn parse_throw_stmt(pair: Pair<Rule>) -> Statement {
    let expr = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::expr_text)
        .map(|c| expr_from_text(c.as_str()));
    Statement::new(StmtKind::Throw {
        expr,
        cause: None,
    })
}

fn parse_break_stmt(pair: Pair<Rule>) -> Statement {
    let level = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::integer)
        .and_then(|c| c.as_str().parse::<u32>().ok());
    Statement::new(StmtKind::Break(match level {
        Some(level) => BreakTarget::Level(level),
        None => BreakTarget::Implicit,
    }))
}

fn parse_continue_stmt(pair: Pair<Rule>) -> Statement {
    let level = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::integer)
        .and_then(|c| c.as_str().parse::<u32>().ok());
    Statement::new(StmtKind::Continue(match level {
        Some(level) => ContinueTarget::Level(level),
        None => ContinueTarget::Implicit,
    }))
}

fn parse_command_statement(pair: Pair<Rule>) -> Statement {
    let text = pair.as_str().trim().to_string();
    if let Some(stmt) = parse_exit_command_statement(&text) {
        return stmt;
    }
    let expr = match pair.as_rule() {
        Rule::command_stmt => parse_command_line(&text)
            .or_else(|| parse_pipeline(pair.clone()))
            .unwrap_or_else(|| expr_from_text(&text)),
        _ => parse_pipeline(pair)
            .or_else(|| parse_command_line(&text))
            .unwrap_or_else(|| expr_from_text(&text)),
    };
    Statement::new(StmtKind::Expr(expr))
}

fn parse_exit_command_statement(text: &str) -> Option<Statement> {
    let tokens = split_command_tokens(text);
    if tokens.is_empty() || !tokens[0].eq_ignore_ascii_case("exit") {
        return None;
    }
    let status = if tokens.len() > 1 {
        Some(expr_from_text(&tokens[1..].join(" ")))
    } else {
        None
    };
    Some(Statement::new(StmtKind::Exit { status }))
}

fn parse_assignment_statement(pair: Pair<Rule>) -> Statement {
    let mut target = None;
    let mut op = "=".to_string();
    let mut rhs = String::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::lvalue => {
                target = Some(assignment_target(child.as_str()));
            }
            Rule::assignment_op => {
                op = child.as_str().to_string();
            }
            Rule::expr_text => {
                rhs = child.as_str().to_string();
            }
            _ => {}
        }
    }

    let target = target.unwrap_or_else(|| Expression::null());
    let value = expr_from_text(rhs.as_str());
    let kind = match op.as_str() {
        "=" => StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        },
        "+=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Add,
            value,
        },
        "-=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Sub,
            value,
        },
        "*=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Mul,
            value,
        },
        "/=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Div,
            value,
        },
        "%=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Mod,
            value,
        },
        "**=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Pow,
            value,
        },
        "&=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitAnd,
            value,
        },
        "|=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitOr,
            value,
        },
        "^=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitXor,
            value,
        },
        "<<=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Shl,
            value,
        },
        ">>=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Shr,
            value,
        },
        "&&=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::And,
            value,
        },
        "||=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Or,
            value,
        },
        "??=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::NullCoalesce,
            value,
        },
        _ => StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(target),
            value: Box::new(value),
        })),
    };

    Statement::new(kind)
}

fn parse_increment_statement(pair: Pair<Rule>) -> Statement {
    let text = pair.as_str();
    let (target_text, is_inc) = text
        .trim_end()
        .strip_suffix("++")
        .map(|target| (target.trim(), true))
        .or_else(|| text.trim_end().strip_suffix("--").map(|target| (target.trim(), false)))
        .unwrap_or(("", false));
    if target_text.is_empty() {
        return Statement::new(StmtKind::Expr(Expression::null()));
    }

    let target = assignment_target(target_text);
    Statement::new(StmtKind::CompoundAssign {
        target,
        op: if is_inc {
            CompoundOp::Add
        } else {
            CompoundOp::Sub
        },
        value: Expression::int(1),
    })
}

fn assignment_target(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    parse_member_expression(text)
}

fn parse_command_line(text: &str) -> Option<Expression> {
    let mut segment_iter = split_command_segments(text).into_iter();
    let first_segment = segment_iter.next()?;
    let mut first_tokens = split_command_tokens(&first_segment);
    if first_tokens.is_empty() {
        return None;
    }

    let (head, mut args) = parse_command_parts(&first_tokens)?;
    let mut expr = Expression::new(ExprKind::Call {
        callee: Box::new(head),
        args,
        optional: false,
    });

    for segment in segment_iter {
        let segment_tokens = split_command_tokens(&segment);
        if segment_tokens.is_empty() {
            continue;
        }
        let (next, mut next_args) = parse_command_parts(&segment_tokens)?;
        let mut chained = vec![Argument::positional(expr)];
        chained.append(&mut next_args);
        expr = Expression::new(ExprKind::Call {
            callee: Box::new(next),
            args: chained,
            optional: false,
        });
    }

    Some(expr)
}

fn detect_assignment_like(text: &str) -> Option<(String, String, String)> {
    let tokens = split_command_tokens(text);
    if tokens.len() >= 3 && tokens[0].starts_with('$') && is_assignment_operator(&tokens[1]) {
        let name = tokens[0].trim_start_matches('$').to_string();
        let rhs = tokens[2..].join(" ");
        let op = tokens[1].clone();
        return Some((name, op, rhs));
    }
    for op in [
        "??=",
        "||=",
        "&&=",
        "<<=",
        ">>=",
        "**=",
        "+=",
        "-=",
        "*=",
        "/=",
        "%=",
        "&=",
        "|=",
        "^=",
        "=",
    ] {
        if let Some((left, right)) = split_once(text, op) {
            let left = left.trim();
            let right = right.trim();
            if left.is_empty() || !left.starts_with('$') {
                continue;
            }
            return Some((left.to_string(), op.to_string(), right.to_string()));
        }
    }
    None
}

fn is_assignment_operator(raw: &str) -> bool {
    matches!(
        raw,
        "=" | "+=" | "-=" | "*=" | "/=" | "%=" | "**=" | "&=" | "|=" | "^=" | "<<=" | ">>=" | "&&=" | "||=" | "??="
    )
}

fn assign_statement(var_name: &str, op: &str, rhs: &str) -> StmtKind {
    let value = expr_from_text(rhs);
    match op {
        "=" => StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(var_name.to_string()),
                type_hint: None,
                init: Some(value),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        "+=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Add,
            value,
        },
        "-=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Sub,
            value,
        },
        "*=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Mul,
            value,
        },
        "/=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Div,
            value,
        },
        "%=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Mod,
            value,
        },
        "**=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Pow,
            value,
        },
        "&=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::BitAnd,
            value,
        },
        "|=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::BitOr,
            value,
        },
        "^=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::BitXor,
            value,
        },
        "<<=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Shl,
            value,
        },
        ">>=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Shr,
            value,
        },
        "&&=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::And,
            value,
        },
        "||=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::Or,
            value,
        },
        "??=" => StmtKind::CompoundAssign {
            target: Expression::ident(var_name),
            op: CompoundOp::NullCoalesce,
            value,
        },
        _ => StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(var_name)),
            value: Box::new(value),
        })),
    }
}

fn parse_pipeline(pair: Pair<Rule>) -> Option<Expression> {
    let mut work = std::collections::VecDeque::new();
    work.push_back(pair);
    let mut segments: Vec<Pair<Rule>> = Vec::new();

    while let Some(node) = work.pop_front() {
        match node.as_rule() {
            Rule::command_stmt | Rule::pipeline_expr => {
                for child in node.into_inner() {
                    work.push_back(child);
                }
            }
            Rule::command_segment => segments.push(node),
            _ => {}
        }
    }

    let mut segments = segments.into_iter();
    let first = segments.next()?;
    let (head, mut args) = parse_command_segment(first)?;
    let mut expr = Expression::new(ExprKind::Call {
        callee: Box::new(head),
        args,
        optional: false,
    });

    for segment in segments {
        let (next, mut next_args) = parse_command_segment(segment)?;
        let mut chained = vec![Argument::positional(expr)];
        chained.append(&mut next_args);
        expr = Expression::new(ExprKind::Call {
            callee: Box::new(next),
            args: chained,
            optional: false,
        });
    }

    Some(expr)
}

fn parse_command_tokens_as_expr(tokens: &[String]) -> Option<Expression> {
    let (callee, args) = parse_command_parts(tokens)?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    }))
}

fn parse_command_segment(pair: Pair<Rule>) -> Option<(Expression, Vec<Argument>)> {
    let mut tokens = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::command_head | Rule::command_arg | Rule::command_token => {
                tokens.push(extract_command_token_text(child));
            }
            _ => {}
        }
    }
    if tokens.is_empty() {
        return None;
    }

    parse_command_parts(&tokens)
}

fn parse_command_parts(tokens: &[String]) -> Option<(Expression, Vec<Argument>)> {
    let callee = parse_command_head(&tokens[0]);
    let mut args = Vec::new();
    let mut i = 1;
    while i < tokens.len() {
        let token = tokens[i].as_str();
        if token.starts_with('-') && token.len() > 1 {
            let flag = &token[1..];
            if let Some((key, value)) = flag.split_once(':') {
                args.push(Argument {
                    value: parse_atom(value),
                    name: Some(key.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 1;
                continue;
            }
            if let Some((key, value)) = flag.split_once('=') {
                args.push(Argument {
                    value: parse_atom(value),
                    name: Some(key.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 1;
                continue;
            }
            if let Some(next) = tokens.get(i + 1) {
                if next.starts_with('-') || next.trim().is_empty() {
                    args.push(Argument {
                        value: Expression::bool(true),
                        name: Some(flag.to_string()),
                        by_ref: false,
                        spread: false,
                    });
                    i += 1;
                    continue;
                }
                args.push(Argument {
                    value: parse_atom(next),
                    name: Some(flag.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 2;
                continue;
            }
            args.push(Argument {
                value: Expression::bool(true),
                name: Some(flag.to_string()),
                by_ref: false,
                spread: false,
            });
            i += 1;
            continue;
        }
        args.push(Argument {
            value: parse_atom(token),
            name: None,
            by_ref: false,
            spread: false,
        });
        i += 1;
    }
    Some((callee, args))
}

fn parse_command_head(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    if looks_like_command_name(text) || is_path_like_command(text) {
        Expression::ident(text)
    } else {
        parse_atom(text)
    }
}

fn is_path_like_command(text: &str) -> bool {
    let text = text.trim();
    text.starts_with('&')
        || text.ends_with(".ps1")
        || text.ends_with(".psm1")
        || text.ends_with(".psd1")
        || text.contains('/')
        || text.contains('\\')
}

fn split_command_segments(input: &str) -> Vec<String> {
    let mut segments = Vec::new();
    let mut current = String::new();
    let mut quote = None;
    let mut depth_paren = 0usize;
    let mut depth_bracket = 0usize;
    let mut depth_brace = 0usize;
    let mut escaped = false;
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.next() {
        if escaped {
            current.push(ch);
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            current.push(ch);
            continue;
        }

        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
        } else if ch == '(' {
            depth_paren += 1;
            current.push(ch);
        } else if ch == ')' {
            depth_paren = depth_paren.saturating_sub(1);
            current.push(ch);
        } else if ch == '[' {
            depth_bracket += 1;
            current.push(ch);
        } else if ch == ']' {
            depth_bracket = depth_bracket.saturating_sub(1);
            current.push(ch);
        } else if ch == '{' {
            depth_brace += 1;
            current.push(ch);
        } else if ch == '}' {
            depth_brace = depth_brace.saturating_sub(1);
            current.push(ch);
        } else if ch == '|' && depth_paren == 0 && depth_bracket == 0 && depth_brace == 0 {
            if !current.trim().is_empty() {
                segments.push(current.trim().to_string());
            } else {
                segments.push(String::new());
            }
            current.clear();
            continue;
        } else if ch == '"' || ch == '\'' {
            quote = Some(ch);
            current.push(ch);
        } else {
            current.push(ch);
        }
    }

    if !current.trim().is_empty() {
        segments.push(current.trim().to_string());
    }

    segments
}

fn extract_command_token_text(pair: Pair<Rule>) -> String {
    let raw = pair.as_str().trim().to_string();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::command_token {
            return child.as_str().trim().to_string();
        }
    }
    raw
}

fn collect_command_tokens(pair: &Pair<Rule>, out: &mut Vec<String>) {
    if pair.as_rule() == Rule::command_token {
        out.push(pair.as_str().trim().to_string());
        return;
    }

    for child in pair.clone().into_inner() {
        collect_command_tokens(&child, out);
    }
}

fn parse_atom(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    if looks_like_command_name(text) && !is_numeric_like(text) {
        return Expression::ident(text);
    }
    expr_from_text(text)
}

fn expr_from_text(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }

    if let Some(s) = strip_surrounded(text, '"') {
        return parse_double_quoted_string(&s);
    }
    if let Some(s) = strip_surrounded(text, '\'') {
        return Expression::string(&s);
    }
    if let Some((type_name, rhs)) = parse_ps_cast(text) {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr_from_text(&rhs)),
            type_name,
        });
    }
    if let Some(expr) = parse_ps_array_literal(text) {
        return expr;
    }
    if let Some(expr) = parse_ps_object_literal(text) {
        return expr;
    }
    if let Some((from, to)) = parse_ps_range(text) {
        return Expression::new(ExprKind::Range {
            start: Box::new(expr_from_text(&from)),
            end: Box::new(expr_from_text(&to)),
            inclusive: true,
        });
    }
    if let Ok(int) = text.parse::<i64>() {
        return Expression::int(int);
    }
    if let Ok(float) = text.parse::<f64>() {
        return Expression::float(float);
    }
    if matches!(text.to_lowercase().as_str(), "$true" | "true") {
        return Expression::bool(true);
    }
    if matches!(text.to_lowercase().as_str(), "$false" | "false") {
        return Expression::bool(false);
    }
    if text == "$null" {
        return Expression::null();
    }

    for (op, bin) in [
        (" -eq ", BinOp::Eq),
        (" -ne ", BinOp::NotEq),
        (" -gt ", BinOp::Gt),
        (" -ge ", BinOp::GtEq),
        (" -lt ", BinOp::Lt),
        (" -le ", BinOp::LtEq),
        (" -and ", BinOp::And),
        (" -or ", BinOp::Or),
        (" -in ", BinOp::In),
        (" -notin ", BinOp::NotIn),
        (" -contains ", BinOp::In),
        (" -notcontains ", BinOp::NotIn),
        (" -like ", BinOp::Like),
        (" -match ", BinOp::Like),
        (" -is ", BinOp::Is),
        (" -isnot ", BinOp::IsNot),
        (" -matchnot ", BinOp::NotIn),
    ] {
        if let Some((left, right)) = split_once(text, op) {
            return Expression::new(ExprKind::Binary {
                op: bin,
                left: Box::new(expr_from_text(left)),
                right: Box::new(expr_from_text(right)),
            });
        }
    }

    if text.starts_with('$') {
        return parse_member_expression(text);
    }

    if is_fully_wrapped(text, '(', ')') {
        return parse_script_expression(text);
    }

    if text.contains('.') || text.contains('[') {
        return parse_member_expression(text);
    }

    if looks_like_command_invocation(text) {
        if let Some(expr) = parse_command_line(text) {
            return expr;
        }
    }

    Expression::ident(text)
}

fn parse_double_quoted_string(text: &str) -> Expression {
    let chars: Vec<char> = text.chars().collect();
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut literal = String::new();
    let mut i = 0;

    while i < chars.len() {
        let ch = chars[i];

        if ch == '`' {
            if let Some(next) = chars.get(i + 1) {
                if let Some(mapped) = parse_powershell_escape(*next) {
                    literal.push(mapped);
                } else {
                    literal.push(*next);
                }
                i += 2;
            } else {
                i += 1;
            }
            continue;
        }

        if ch != '$' {
            literal.push(ch);
            i += 1;
            continue;
        }

        if i + 1 >= chars.len() {
            literal.push(ch);
            i += 1;
            continue;
        }

        if !literal.is_empty() {
            parts.push(InterpolPart::Text(std::mem::take(&mut literal)));
            literal.clear();
        }

        let next = chars[i + 1];
        match next {
            '$' | '`' => {
                literal.push('$');
                i += 2;
            }

            '{' => {
                if let Some(close) = find_matching_in_chars(&chars, i + 1, '{', '}') {
                    let inner = chars[(i + 2)..close].iter().collect::<String>();
                    parts.push(InterpolPart::Expr(parse_script_expression(&inner)));
                    i = close + 1;
                } else {
                    literal.push('$');
                    literal.push('{');
                    i += 2;
                }
            }

            '(' => {
                if let Some(close) = find_matching_in_chars(&chars, i + 1, '(', ')') {
                    let inner = chars[(i + 2)..close].iter().collect::<String>();
                    parts.push(InterpolPart::Expr(parse_script_expression(&inner)));
                    i = close + 1;
                } else {
                    literal.push('$');
                    i += 1;
                }
            }

            c if is_ps_variable_char(c) => {
                let mut end = i + 1;
                while end < chars.len() && is_ps_variable_char(chars[end]) {
                    end += 1;
                }

                let name = chars[(i + 1)..end].iter().collect::<String>();
                parts.push(InterpolPart::Expr(parse_script_expression(&format!("${}", name))));
                i = end;
            }

            _ => {
                literal.push('$');
                i += 1;
            }
        }
    }

    if !literal.is_empty() {
        parts.push(InterpolPart::Text(std::mem::take(&mut literal)));
    }

    match parts.as_slice() {
        [] => Expression::string(""),
        [InterpolPart::Text(s)] => Expression::string(s),
        [InterpolPart::Expr(expr)] => expr.clone(),
        _ => {
            Expression::new(ExprKind::Interpolation(parts))
        }
    }
}

fn parse_script_expression(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }

    let text = strip_outer_parentheses(text);
    if text.is_empty() {
        return Expression::null();
    }

    if has_binary_operator(text) {
        return expr_from_text(text);
    }

    let mut parts = split_with_depth(text, '.', 0)
        .into_iter()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();

    if parts.is_empty() {
        return Expression::ident(text);
    }

    let mut expr = parse_script_head(parts.remove(0));
    for part in parts {
        expr = apply_script_segment(expr, &part);
    }

    expr
}

fn has_binary_operator(text: &str) -> bool {
    for (op, _) in [
        (" -eq ", BinOp::Eq),
        (" -ne ", BinOp::NotEq),
        (" -gt ", BinOp::Gt),
        (" -ge ", BinOp::GtEq),
        (" -lt ", BinOp::Lt),
        (" -le ", BinOp::LtEq),
        (" -and ", BinOp::And),
        (" -or ", BinOp::Or),
        (" -in ", BinOp::In),
        (" -notin ", BinOp::NotIn),
        (" -contains ", BinOp::In),
        (" -notcontains ", BinOp::NotIn),
        (" -like ", BinOp::Like),
        (" -matchnot ", BinOp::NotIn),
        (" -match ", BinOp::Like),
        (" -is ", BinOp::Is),
        (" -isnot ", BinOp::IsNot),
    ] {
        if text.contains(op) {
            return true;
        }
    }
    false
}

fn parse_script_head(raw: String) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }

    if text.starts_with('$') {
        return parse_member_expression(text);
    }

    if is_fully_wrapped(text, '(', ')') {
        return parse_script_expression(&text[1..text.len() - 1]);
    }

    if looks_like_command_invocation(text) {
        if let Some(expr) = parse_command_line(text) {
            return expr;
        }
    }

    expr_from_text(text)
}

fn apply_script_segment(object: Expression, segment: &str) -> Expression {
    if segment.is_empty() {
        return object;
    }

    if let Some((name, args)) = parse_call_segment(segment) {
        let callee = Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        });
        let args = parse_call_args(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        });
    }

    if let Some((_, tail)) = segment.split_once('[') {
        if let Some(close) = tail.rfind(']') {
            let index = expr_from_text(&tail[..close]);
            return Expression::new(ExprKind::Index {
                object: Box::new(object),
                index: Box::new(index),
                null_safe: false,
            });
        }
    }

    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: segment.to_string(),
        null_safe: false,
    })
}

fn parse_call_segment(segment: &str) -> Option<(&str, &str)> {
    if segment.is_empty() || !segment.ends_with(')') {
        return None;
    }

    let open = segment.find('(')?;
    if open == 0 || open == segment.len() - 1 {
        return None;
    }

    let chars: Vec<char> = segment.chars().collect();
    let open_idx = segment[..open].chars().count();
    let close = find_matching_in_chars(&chars, open_idx, '(', ')')?;
    if close + 1 != chars.len() {
        return None;
    }

    let name = segment[..open].trim();
    if name.is_empty() {
        return None;
    }

    Some((name, &segment[open + 1..segment.len() - 1]))
}

fn parse_call_args(raw: &str) -> Vec<Argument> {
    if raw.trim().is_empty() {
        return Vec::new();
    }

    split_top_level(raw, ',')
        .into_iter()
        .filter(|s| !s.trim().is_empty())
        .map(|arg| Argument::positional(expr_from_text(arg.trim())))
        .collect()
}

fn is_numeric_like(raw: &str) -> bool {
    let text = raw.trim();
    if text.is_empty() {
        return false;
    }

    text.parse::<i64>().is_ok() || text.parse::<f64>().is_ok() || text == "$null"
}

fn parse_ps_cast(text: &str) -> Option<(String, String)> {
    let trimmed = text.trim();
    if !trimmed.starts_with('[') {
        return None;
    }

    let close = trimmed.find(']')?;
    let type_name = trimmed[1..close].trim();
    if type_name.is_empty() {
        return None;
    }

    let rhs = trimmed[close + 1..].trim();
    if rhs.is_empty() {
        return None;
    }

    Some((type_name.to_string(), rhs.to_string()))
}

fn parse_ps_array_literal(text: &str) -> Option<Expression> {
    let trimmed = text.trim();
    if !(trimmed.starts_with("@(") && trimmed.ends_with(')')) {
        return None;
    }

    let inner = trimmed[2..trimmed.len() - 1].trim();
    if inner.is_empty() {
        return Some(Expression::new(ExprKind::Array(Vec::new())));
    }

    let elements = split_top_level(inner, ',')
        .into_iter()
        .filter(|s| !s.trim().is_empty())
        .map(|value| ArrayElement {
            key: None,
            value: expr_from_text(value.trim()),
            spread: false,
            by_ref: false,
        })
        .collect();

    Some(Expression::new(ExprKind::Array(elements)))
}

fn parse_ps_object_literal(text: &str) -> Option<Expression> {
    let trimmed = text.trim();
    if !(trimmed.starts_with("@{") && trimmed.ends_with('}')) {
        return None;
    }

    let inner = trimmed[2..trimmed.len() - 1].trim();
    if inner.is_empty() {
        return Some(Expression::new(ExprKind::Object(Vec::new())));
    }

    let properties = split_object_fields(inner)
        .into_iter()
        .filter_map(|item| {
            let mut parts = item.splitn(2, '=').collect::<Vec<_>>();
            if parts.len() != 2 {
                parts = item.splitn(2, ':').collect();
                if parts.len() != 2 {
                    return None;
                }
            }

            let key = parts[0].trim();
            let value = parts[1].trim();
            if key.is_empty() || value.is_empty() {
                return None;
            }

            Some(ObjectProperty::KeyValue {
                key: Expression::string(key),
                value: expr_from_text(value),
            })
        })
        .collect::<Vec<_>>();

    Some(Expression::new(ExprKind::Object(properties)))
}

fn parse_ps_range(text: &str) -> Option<(String, String)> {
    let trimmed = text.trim();
    let (left, right) = split_once(trimmed, "..")?;
    if left.is_empty() || right.is_empty() {
        return None;
    }
    Some((left.to_string(), right.to_string()))
}

fn looks_like_command_invocation(raw: &str) -> bool {
    let text = raw.trim();
    if text.is_empty() || is_numeric_like(text) {
        return false;
    }

    if text.contains(' ') {
        if let Some(head) = text.split_whitespace().next() {
            return looks_like_command_name(head) || is_path_like_command(head);
        }
        return false;
    }

    looks_like_command_name(text) || is_path_like_command(text)
}

fn parse_powershell_escape(next: char) -> Option<char> {
    match next {
        '`' => Some('`'),
        '"' => Some('"'),
        '\'' => Some('\''),
        '$' => Some('$'),
        'n' => Some('\n'),
        'r' => Some('\r'),
        't' => Some('\t'),
        '0' => Some('\0'),
        _ => None,
    }
}

fn split_object_fields(text: &str) -> Vec<String> {
    let mut fields = Vec::new();
    let mut start = 0usize;
    let mut depth = 0usize;
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;

    for (idx, ch) in text.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            continue;
        }

        if in_single {
            if ch == '\'' {
                in_single = false;
            }
            continue;
        }

        if in_double {
            if ch == '"' {
                in_double = false;
            }
            continue;
        }

        match ch {
            '\'' => in_single = true,
            '"' => in_double = true,
            '(' | '{' | '[' => depth += 1,
            ')' | '}' | ']' => depth = depth.saturating_sub(1),
            ',' | ';' if depth == 0 => {
                fields.push(text[start..idx].trim().to_string());
                start = idx + ch.len_utf8();
            }
            _ => {}
        }
    }

    if start <= text.len() {
        fields.push(text[start..].trim().to_string());
    }

    fields.into_iter().filter(|f| !f.is_empty()).collect()
}

fn looks_like_command_name(raw: &str) -> bool {
    let mut chars = raw.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }

    chars.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-')
}

fn is_ps_variable_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '-' || c == ':'
}

fn is_fully_wrapped(text: &str, open: char, close: char) -> bool {
    if text.is_empty() || !text.starts_with(open) || !text.ends_with(close) {
        return false;
    }

    let chars: Vec<char> = text.chars().collect();
    find_matching_in_chars(&chars, 0, open, close)
        .is_some_and(|idx| idx == chars.len() - 1)
}

fn strip_outer_parentheses(text: &str) -> &str {
    let mut out = text.trim();
    while is_fully_wrapped(out, '(', ')') {
        out = out[1..out.len() - 1].trim();
    }
    out
}

fn find_matching_in_chars(chars: &[char], open_idx: usize, open: char, close: char) -> Option<usize> {
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;
    let mut depth = 0usize;

    for (idx, ch) in chars.iter().enumerate() {
        let ch = *ch;

        if escaped {
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            continue;
        }

        if in_single {
            if ch == '\'' {
                in_single = false;
            }
            continue;
        }

        if in_double {
            if ch == '"' {
                in_double = false;
                continue;
            }

            if ch == open {
                depth += 1;
            } else if ch == close && depth > 0 {
                depth -= 1;
                if depth == 0 {
                    return Some(idx);
                }
            }

            continue;
        }

        if ch == '\'' {
            in_single = true;
            continue;
        }

        if ch == '"' {
            in_double = true;
            continue;
        }

        if idx < open_idx {
            continue;
        }

        if idx == open_idx {
            if ch == open {
                depth = 1;
            }
            continue;
        }

        if ch == open {
            depth += 1;
        } else if ch == close {
            if depth == 0 {
                return None;
            }
            depth -= 1;
            if depth == 0 {
                return Some(idx);
            }
        }
    }

    None
}

fn split_top_level(text: &str, sep: char) -> Vec<String> {
    let mut parts: Vec<String> = Vec::new();
    let mut depth = 0isize;
    let mut start = 0usize;
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;

    for (idx, ch) in text.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            continue;
        }

        if in_single {
            if ch == '\'' {
                in_single = false;
            }
            continue;
        }

        if in_double {
            if ch == '"' {
                in_double = false;
                continue;
            }
            match ch {
                '(' | '[' => depth += 1,
                ')' | ']' => depth -= 1,
                _ => {}
            }
            continue;
        }

        if ch == '\'' {
            in_single = true;
            continue;
        }

        if ch == '"' {
            in_double = true;
            continue;
        }

        match ch {
            '(' | '[' => depth += 1,
            ')' | ']' => {
                if depth > 0 {
                    depth -= 1;
                }
            }
            _ => {}
        }

        if ch == sep && depth == 0 {
            parts.push(text[start..idx].to_string());
            start = idx + ch.len_utf8();
        }
    }

    parts.push(text[start..].to_string());
    parts
}

fn parse_member_expression(raw: &str) -> Expression {
    let mut pieces = split_with_depth(raw, '.', 0)
        .into_iter()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();

    if pieces.is_empty() {
        return Expression::ident(raw);
    }

    let root = pieces.remove(0);
    let base = if root.starts_with('$') {
        Expression::ident(root.trim_start_matches('$'))
    } else {
        Expression::ident(&root)
    };

    pieces.into_iter().fold(base, |expr, field| {
        if let Some((_, tail)) = field.split_once('[') {
            if let Some(close) = tail.rfind(']') {
                let index = expr_from_text(&tail[..close]);
                Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
                    null_safe: false,
                })
            } else {
                Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: field,
                    null_safe: false,
                })
            }
        } else {
            Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field,
                null_safe: false,
            })
        }
    })
}

fn split_once<'a>(input: &'a str, op: &'a str) -> Option<(&'a str, &'a str)> {
    input.find(op).map(|idx| {
        let left = input[..idx].trim();
        let right = input[idx + op.len()..].trim();
        (left, right)
    })
}

fn strip_surrounded(text: &str, quote: char) -> Option<String> {
    if text.len() >= 2 && text.starts_with(quote) && text.ends_with(quote) {
        Some(text[1..text.len() - 1].to_string())
    } else {
        None
    }
}

fn split_with_depth(text: &str, sep: char, start_depth: usize) -> Vec<String> {
    let mut out = Vec::new();
    let mut depth = start_depth;
    let mut start = 0;
    let mut last = 0usize;
    for (i, ch) in text.char_indices() {
        match ch {
            '[' | '(' => depth += 1,
            ']' | ')' => depth = depth.saturating_sub(1),
            c if c == sep && depth == 0 => {
                out.push(text[start..i].to_string());
                start = i + 1;
            }
            _ => {}
        }
        last = i + ch.len_utf8();
    }
    if start <= last {
        out.push(text[start..].to_string());
    }
    out
}

fn split_command_tokens(input: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut quote = None;
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '\\' {
            if let Some(next) = chars.next() {
                current.push(next);
            }
            continue;
        }
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        } else if ch == '"' || ch == '\'' {
            quote = Some(ch);
            current.push(ch);
            continue;
        }

        if ch.is_whitespace() && quote.is_none() {
            if !current.trim().is_empty() {
                tokens.push(current.trim().to_string());
                current.clear();
            }
            continue;
        }
        current.push(ch);
    }

    if !current.trim().is_empty() {
        tokens.push(current.trim().to_string());
    }

    tokens
}

trait IdentEmpty {
    fn as_ident_empty(&self) -> bool;
}

impl IdentEmpty for Expression {
    fn as_ident_empty(&self) -> bool {
        match &self.kind {
            ExprKind::Ident(v) => v.is_empty(),
            _ => false,
        }
    }
}
