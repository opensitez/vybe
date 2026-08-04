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
            // `[CmdletBinding()]`, `[OutputType(…)]` — declaration metadata that
            // carries no runtime behaviour.
            Rule::attribute_stmt => return Ok(None),
            Rule::using_stmt => return Ok(None),
            Rule::assignment_stmt => return Ok(Some(parse_assignment_statement(pair))),
            Rule::increment_stmt => return Ok(Some(parse_increment_statement(pair))),
            Rule::expr_stmt => {
                let expr = pair
                    .into_inner()
                    .next()
                    .map(walk_expr)
                    .unwrap_or_else(Expression::null);
                return Ok(Some(Statement::new(StmtKind::Expr(expr))));
            }
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
    if pair.as_rule() == Rule::expression {
        return Ok(Some(Statement::new(StmtKind::Expr(walk_expr(pair)))));
    }
    let _ = body;
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
            // Walk the DOTTED_NAME children, not the rule's raw text — the text
            // still carries the leading `:` and would yield a parent named
            // ": A" that resolves to nothing.
            Rule::class_heritage => {
                for base in child.into_inner() {
                    let name = base.as_str().trim();
                    if !name.is_empty() {
                        parents.push(name.to_string());
                    }
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
        if let Some(member) = parse_class_member(child)? {
            members.push(member);
        }
    }
    Ok(members)
}

/// `[string] Speak() { … }` — a method, identified by its declared return type.
fn parse_ps_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::class_modifier => {
                if child.as_str().eq_ignore_ascii_case("static") {
                    is_static = true;
                }
            }
            Rule::member_name => name = child.as_str().to_string(),
            Rule::function_params => params = parse_function_params(child),
            Rule::block => {
                body = implicit_return(parse_block_with_function_params(child, &mut params)?)
            }
            _ => {}
        }
    }

    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            body,
            modifiers,
            is_async: false,
            is_generator: false,
            is_sub: false,
            return_type: None,
            handles: Vec::new(),
        },
    ))))
}

/// `Animal([string]$name) { … }` — a constructor: class-named, no return type.
fn parse_ps_constructor(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::IDENT => name = Some(child.as_str().to_string()),
            Rule::function_params => params = parse_function_params(child),
            Rule::ctor_base => {
                let args = child
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::arg_list)
                    .map(walk_arg_list)
                    .unwrap_or_default();
                base_args = Some(args);
            }
            Rule::block => body = parse_block_with_function_params(child, &mut params)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name,
        params,
        body,
        base_args,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

/// `[string]$Name = 'x'` — a field, with an optional type and initialiser.
fn parse_ps_property(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut type_hint = None;
    let mut name = String::new();
    let mut init = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::class_modifier => {
                if child.as_str().eq_ignore_ascii_case("static") {
                    is_static = true;
                }
            }
            Rule::type_literal => {
                let t = type_literal_name(child.as_str()).trim();
                if !t.is_empty() {
                    type_hint = Some(t.to_string());
                }
            }
            Rule::var_ref => {
                name = scope_qualified_name(child.as_str().trim_start_matches('$')).to_string();
            }
            _ => init = Some(walk_expr(child)),
        }
    }

    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;

    Ok(ClassMember::Field {
        name,
        type_hint,
        init,
        modifiers,
        with_events: false,
        array_bounds: None,
    })
}

fn parse_class_member(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::ps_method => parse_ps_method(pair).map(Some),
        Rule::ps_constructor => parse_ps_constructor(pair).map(Some),
        Rule::ps_property => parse_ps_property(pair).map(Some),
        Rule::class_function_decl => {
            let statement = parse_function_decl(pair)?;
            Ok(Some(ClassMember::Method(Box::new(statement))))
        }
        Rule::constructor_decl => parse_constructor_decl(pair).map(Some),
        _ => Ok(None),
    }
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
            Rule::block => {
                body = implicit_return(parse_block_with_function_params(child, &mut params)?)
            }
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
                Rule::type_literal => {
                    let inner = type_literal_name(piece.as_str()).trim();
                    if !inner.is_empty() {
                        type_hint = Some(inner.to_string());
                    }
                }
                Rule::var_ref | Rule::IDENT => {
                    name = scope_qualified_name(piece.as_str().trim_start_matches('$')).to_string();
                }
                Rule::expression => {
                    default = Some(walk_expr(piece));
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
    // Hand `parse_function_params` the ENCLOSING pair, not the
    // `function_param_list` itself: it descends two levels, so passing the list
    // skipped straight past every `function_param`.
    out.extend(parse_function_params(pair));
}

fn parse_block_with_function_params(
    pair: Pair<Rule>,
    params: &mut Vec<Param>,
) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();

    // `statement` is a SILENT rule, so a block's children are the concrete
    // statement rules (`assignment_stmt`, `command_stmt`, …) and never
    // `Rule::statement` itself. Matching on that name here skipped every
    // statement and left function and constructor bodies empty.
    for child in pair.into_inner() {
        if child.as_rule() == Rule::param_stmt {
            parse_param_stmt(child, params);
            continue;
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
            Rule::condition_expr => cond = Some(walk_condition(child)),
            Rule::block => then_body = parse_block_statements(child)?,
            Rule::elseif_stmt => {
                let mut branch_cond = None;
                let mut branch_body = Vec::new();
                for part in child.into_inner() {
                    if part.as_rule() == Rule::condition_expr {
                        branch_cond = Some(walk_condition(part));
                    } else if part.as_rule() == Rule::block {
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

/// `condition_expr` wraps a parenthesised expression — unwrap to the inner
/// `expression` pair rather than re-parsing the text.
fn walk_condition(pair: Pair<Rule>) -> Expression {
    pair.into_inner()
        .find(|c| matches!(c.as_rule(), Rule::expression | Rule::command_pipeline))
        .map(walk_expr)
        .unwrap_or_else(Expression::null)
}

fn parse_foreach_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut iter_var = String::new();
    let mut iter = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::var_ref => {
                iter_var = scope_qualified_name(child.as_str().trim_start_matches('$')).to_string()
            }
            Rule::expression | Rule::command_pipeline => iter = Some(walk_expr(child)),
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
            Rule::expression => {
                expr = Some(walk_expr(child));
            }
            Rule::switch_case => {
                let mut branch_body = None;
                let mut branch_conditions = Vec::new();
                let is_default = child
                    .as_str()
                    .trim_start()
                    .to_lowercase()
                    .starts_with("default");

                for part in child.into_inner() {
                    match part.as_rule() {
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
                                    Rule::expression => {
                                        branch_conditions
                                            .push(CaseCondition::Value(walk_expr(inner)));
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
    let mut init = None;
    let mut cond = None;
    let mut update = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::for_init => {
                if let Some(inner) = child.into_inner().next() {
                    init = parse_statement(inner)?;
                }
            }
            Rule::for_cond => cond = Some(walk_expr(child)),
            Rule::for_update => {
                if let Some(inner) = child.into_inner().next() {
                    update = parse_statement(inner)?;
                }
            }
            Rule::block => body = parse_block_statements(child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::For {
        init: init.map(Box::new),
        cond,
        update: update.map(statement_to_expression),
        body,
    }))
}

/// `StmtKind::For.update` is an `Expression`, but the update slot parses as a
/// statement (`$i++`, `$i = $i + 1`). Re-shape assignment statements into the
/// equivalent assignment expressions the shared loop lowering expects.
fn statement_to_expression(stmt: Statement) -> Expression {
    match stmt.kind {
        StmtKind::Expr(expr) => expr,
        StmtKind::Assign {
            mut targets, value, ..
        } => {
            let target = if targets.is_empty() {
                Expression::null()
            } else {
                targets.remove(0)
            };
            Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            })
        }
        StmtKind::CompoundAssign { target, op, value } => {
            let bin = compound_to_binop(op);
            binary_assign(target, bin, value)
        }
        _ => Expression::null(),
    }
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        _ => BinOp::Add,
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
        match child.as_rule() {
            Rule::condition_expr => cond = Some(walk_condition(child)),
            Rule::block => body = parse_block_statements(child)?,
            _ => {}
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
            Rule::do_while_while | Rule::do_while_until => {
                until = child.as_rule() == Rule::do_while_until;
                for inner in child.into_inner() {
                    if inner.as_rule() == Rule::condition_expr {
                        cond = Some(walk_condition(inner));
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
        .find(|c| matches!(c.as_rule(), Rule::expression | Rule::command_pipeline))
        .map(walk_expr);
    Statement::new(StmtKind::Return(expr))
}

fn parse_throw_stmt(pair: Pair<Rule>) -> Statement {
    let expr = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::expression)
        .map(walk_expr);
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
    // Prefer the PEST TREE over the text tokenizer. The text path splits on
    // spaces without tracking parens, so `Write-Host (I 5)` came apart into
    // `(I` and `5)`; the grammar already groups that correctly.
    let expr = parse_pipeline(pair.clone())
        .or_else(|| parse_command_line(&text))
        .unwrap_or_else(|| expr_from_text(&text));
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
    let mut rhs = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::lvalue => {
                target = Some(walk_lvalue(child));
            }
            Rule::assignment_op => {
                op = child.as_str().to_string();
            }
            // `rhs_value` is a silent rule: the RHS arrives as either an
            // `expression` or a `command_pipeline` (`$x = Get-Item | …`).
            Rule::expression | Rule::command_pipeline => {
                rhs = Some(walk_expr(child));
            }
            _ => {}
        }
    }

    let target = target.unwrap_or_else(Expression::null);
    let value = rhs.unwrap_or_else(Expression::null);
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
    let mut target = None;
    let mut is_inc = true;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::lvalue => target = Some(walk_lvalue(child)),
            Rule::increment_op => is_inc = child.as_str() == "++",
            _ => {}
        }
    }

    let Some(target) = target else {
        return Statement::new(StmtKind::Expr(Expression::null()));
    };

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

/// An assignment target: a variable, optionally followed by `.member` /
/// `[index]` steps. Same node shapes the expression walker produces.
fn walk_lvalue(pair: Pair<Rule>) -> Expression {
    // A leading `[int]` is a type constraint on the target, not part of its
    // identity — the shared compiler infers from the assigned value.
    let mut inner = pair.into_inner().peekable();
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_literal) {
        inner.next();
    }

    let mut expr = match inner.next() {
        Some(p) if p.as_rule() == Rule::var_ref => walk_var_ref(p.as_str()),
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };

    for step in inner {
        match step.as_rule() {
            Rule::member_get => {
                let name = step
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().trim_start_matches('$').to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
            }
            Rule::index_get => {
                let index = step
                    .into_inner()
                    .next()
                    .map(walk_expr)
                    .unwrap_or_else(Expression::null);
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
                    null_safe: false,
                });
            }
            _ => {}
        }
    }

    expr
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
        return Expression::ident(text);
    }
    // A bare word in HEAD position is a command name, not the string argument
    // `parse_atom` would produce — a string callee can never be callable.
    if is_bare_command_word(text) {
        return Expression::ident(text);
    }
    parse_atom(text)
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

// ────────────────────────────────────────────────────────────────────────────
// Expression walking
//
// The pest expression grammar is the ONE place operator precedence is defined.
// Nothing below re-splits source text on operator spellings — a text fragment
// captured in command-token position is re-parsed through `Rule::expr_entry`,
// i.e. through the same grammar, so there is a single source of truth.
// ────────────────────────────────────────────────────────────────────────────

/// Parse a text fragment in **expression mode** — a bare word is an identifier.
fn expr_from_text(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    parse_expr_fragment(text).unwrap_or_else(|| Expression::ident(text))
}

/// Parse a token in **command mode** — a bare word is a string argument, the
/// way PowerShell treats `Write-Host FAIL`.
fn parse_atom(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    if is_bare_command_word(text) {
        return Expression::string(text);
    }
    parse_expr_fragment(text).unwrap_or_else(|| Expression::string(text))
}

/// A token that carries no expression syntax at all — in command mode this is
/// a literal string argument.
fn is_bare_command_word(text: &str) -> bool {
    !text.is_empty()
        && !text.starts_with('$')
        && !text.starts_with('@')
        && !text.starts_with('[')
        && !text.starts_with('(')
        && !text.starts_with('"')
        && !text.starts_with('\'')
        && !text.starts_with('&')
        && !is_numeric_like(text)
        && text
            .chars()
            .all(|c| c.is_alphanumeric() || matches!(c, '_' | '-' | '.' | ':' | '\\' | '/' | '*'))
        && text.chars().any(|c| c.is_alphabetic())
}

fn parse_expr_fragment(text: &str) -> Option<Expression> {
    let pairs = super::PowerShellParser::parse(Rule::expr_entry, text).ok()?;
    let root = pairs.into_iter().next()?;
    let expr = root
        .into_inner()
        .find(|c| c.as_rule() == Rule::expression)?;
    Some(walk_expr(expr))
}

/// Walk one expression pair into the shared AST.
fn walk_expr(pair: Pair<Rule>) -> Expression {
    match pair.as_rule() {
        // `(Get-Date)` is a command INVOCATION, not a read of a name. Only a
        // lone bare word qualifies — `($x)` and `(1 + 2)` stay expressions.
        Rule::paren_expr => match lone_bare_word(&pair) {
            Some(name) => Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false,
            }),
            None => first_inner_expr(pair),
        },

        Rule::expression | Rule::expr_statement | Rule::for_cond | Rule::condition_expr => {
            first_inner_expr(pair)
        }

        Rule::comma_expr => walk_comma_expr(pair),
        Rule::ternary_expr => walk_ternary(pair),

        Rule::logical_or
        | Rule::logical_and
        | Rule::comparison
        | Rule::additive
        | Rule::multiplicative
        | Rule::power => walk_binary_chain(pair),

        Rule::range_expr => walk_range(pair),
        Rule::unary => walk_unary(pair),
        Rule::cast_expr => walk_cast(pair),
        Rule::postfix => walk_postfix(pair),

        Rule::command_pipeline => walk_command_pipeline(pair),
        Rule::command_segment => walk_command_segment_expr(pair),
        Rule::array_expr => walk_array_expr(pair),
        Rule::hash_literal => walk_hash_literal(pair),
        Rule::sub_expr => walk_sub_expr(pair),
        Rule::script_block_expr => walk_script_block_expr(pair),

        Rule::number => walk_number(pair.as_str()),
        Rule::quoted_string => {
            let raw = pair.as_str();
            parse_double_quoted_string(raw.get(1..raw.len().saturating_sub(1)).unwrap_or(""))
        }
        Rule::single_quoted_string => {
            let raw = pair.as_str();
            let inner = raw.get(1..raw.len().saturating_sub(1)).unwrap_or("");
            Expression::string(&inner.replace("''", "'"))
        }
        Rule::var_ref => walk_var_ref(pair.as_str()),
        Rule::type_literal => type_literal_expr(type_literal_name(pair.as_str())),
        Rule::bare_word => walk_bare_word(pair.as_str()),

        _ => first_inner_expr(pair),
    }
}

/// The single `bare_word` a group bottoms out in, if the group contains nothing
/// else — no operators, no arguments, no member access. Walks the single-child
/// chain and gives up the moment a level has siblings.
fn lone_bare_word(pair: &Pair<Rule>) -> Option<String> {
    let mut current = pair.clone();
    loop {
        if current.as_rule() == Rule::bare_word {
            return Some(current.as_str().to_string());
        }
        let mut inner = current.clone().into_inner();
        let only = inner.next()?;
        if inner.next().is_some() {
            return None;
        }
        current = only;
    }
}

fn first_inner_expr(pair: Pair<Rule>) -> Expression {
    pair.into_inner()
        .next()
        .map(walk_expr)
        .unwrap_or_else(Expression::null)
}

/// `1,2,3` builds an array; a single element passes straight through.
fn walk_comma_expr(pair: Pair<Rule>) -> Expression {
    let mut items: Vec<Expression> = pair.into_inner().map(walk_expr).collect();
    match items.len() {
        0 => Expression::null(),
        1 => items.remove(0),
        _ => Expression::new(ExprKind::Array(
            items
                .into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )),
    }
}

fn walk_ternary(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let cond = match inner.next() {
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };
    let Some(then_pair) = inner.next() else {
        return cond;
    };
    let then_expr = walk_expr(then_pair);
    let else_expr = inner.next().map(walk_expr).unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(else_expr),
    })
}

/// Left-associative fold over `operand (op operand)*`.
fn walk_binary_chain(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let mut left = match inner.next() {
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };
    while let Some(op_pair) = inner.next() {
        let Some(rhs_pair) = inner.next() else { break };
        let right = walk_expr(rhs_pair);
        left = build_binary(op_pair.as_str(), left, right);
    }
    left
}

fn walk_range(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let start = match inner.next() {
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };
    match inner.next() {
        Some(end) => Expression::new(ExprKind::Range {
            start: Box::new(start),
            end: Box::new(walk_expr(end)),
            inclusive: true,
        }),
        None => start,
    }
}

fn walk_unary(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let Some(first) = inner.next() else {
        return Expression::null();
    };
    if first.as_rule() != Rule::unary_op {
        return walk_expr(first);
    }
    let op_text = first.as_str().trim().to_lowercase();
    let operand = inner.next().map(walk_expr).unwrap_or_else(Expression::null);

    // `-join` / `-split` in unary position are PowerShell's collection forms;
    // keep them as ordinary method calls so shared dispatch owns them.
    match op_text.as_str() {
        "-join" => return method_call_expr(operand, "join", vec![Expression::string("")]),
        "-split" => return method_call_expr(operand, "split", vec![Expression::string(" ")]),
        _ => {}
    }

    let op = match op_text.as_str() {
        "!" | "-not" => UnaryOp::Not,
        "-bnot" => UnaryOp::BitNot,
        "+" => UnaryOp::Pos,
        _ => UnaryOp::Neg,
    };
    Expression::new(ExprKind::Unary {
        op,
        expr: Box::new(operand),
    })
}

fn walk_cast(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let type_name = inner
        .next()
        .map(|p| type_literal_name(p.as_str()).to_string())
        .unwrap_or_default();
    let expr = inner.next().map(walk_expr).unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name,
    })
}

fn walk_postfix(pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let mut expr = match inner.next() {
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };

    for op in inner {
        match op.as_rule() {
            Rule::member_get => {
                let name = op
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().trim_start_matches('$').to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: normalize_member_name(&name),
                    null_safe: false,
                });
            }
            Rule::static_member => {
                let name = op
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
            }
            Rule::method_call | Rule::static_call => {
                let is_static = op.as_rule() == Rule::static_call;
                let mut parts = op.into_inner();
                let name = parts
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                let args = parts.next().map(walk_arg_list).unwrap_or_default();

                // `[Animal]::new(…)` is construction, not a static call — hand
                // the shared compiler the `New` node it already knows.
                if is_static && name.eq_ignore_ascii_case("new") {
                    if let Some(class) = type_name_of(&expr) {
                        expr = Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident(&class)),
                            args: args
                                .into_iter()
                                .map(|value| Argument {
                                    value,
                                    name: None,
                                    by_ref: false,
                                    spread: false,
                                })
                                .collect(),
                        });
                        continue;
                    }
                }

                expr = method_call_expr(expr, &name, args);
            }
            Rule::index_get => {
                let index = op
                    .into_inner()
                    .next()
                    .map(walk_expr)
                    .unwrap_or_else(Expression::null);
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
                    null_safe: false,
                });
            }
            _ => {}
        }
    }

    expr
}

fn walk_arg_list(pair: Pair<Rule>) -> Vec<Expression> {
    pair.into_inner().map(walk_expr).collect()
}

fn method_call_expr(receiver: Expression, name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: name.to_string(),
            null_safe: false,
        })),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

/// `a | b | c` — each stage's value becomes argument 0 of the next callee, so
/// the shared call resolution in `primitives/calls.rs` keeps full control.
fn walk_command_pipeline(pair: Pair<Rule>) -> Expression {
    let mut stages = pair.into_inner();
    let mut expr = match stages.next() {
        Some(p) => walk_expr(p),
        None => return Expression::null(),
    };

    for stage in stages {
        let (callee, mut args) = match stage.as_rule() {
            Rule::command_segment => match parse_command_segment(stage) {
                Some(parts) => parts,
                None => continue,
            },
            _ => (walk_expr(stage), Vec::new()),
        };

        // The core pipeline cmdlets are collection operations. Rewriting them
        // to the equivalent method call lets the existing `[array_methods]`
        // dispatch handle them — no cmdlet builtins, no emitter arm.
        if let ExprKind::Ident(name) = &callee.kind {
            if let Some(method) = pipeline_cmdlet_method(name) {
                let positional: Vec<Expression> = args
                    .iter()
                    .filter(|a| a.name.is_none())
                    .map(|a| a.value.clone())
                    .collect();
                expr = match method {
                    // `… | Out-Null` discards the pipeline value.
                    "" => Expression::null(),
                    m => method_call_expr(expr, m, positional),
                };
                continue;
            }
        }
        args.insert(
            0,
            Argument {
                value: expr,
                name: None,
                by_ref: false,
                spread: false,
            },
        );
        expr = Expression::new(ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        });
    }

    expr
}

/// The method a pipeline cmdlet is equivalent to, including its `?` / `%`
/// aliases. `""` means the stage drops the value (`Out-Null`).
fn pipeline_cmdlet_method(name: &str) -> Option<&'static str> {
    match name.to_lowercase().as_str() {
        "where-object" | "where" | "?" => Some("Where"),
        "foreach-object" | "foreach" | "%" => Some("ForEach"),
        "sort-object" | "sort" => Some("Sort"),
        "measure-object" | "measure" => Some("Count"),
        "out-null" => Some(""),
        _ => None,
    }
}

/// A bare command invocation used where an expression is expected: `(hi 'PASS')`.
fn walk_command_segment_expr(pair: Pair<Rule>) -> Expression {
    match parse_command_segment(pair) {
        Some((callee, args)) => Expression::new(ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        }),
        None => Expression::null(),
    }
}

fn walk_array_expr(pair: Pair<Rule>) -> Expression {
    Expression::new(ExprKind::Array(
        pair.into_inner()
            .map(|p| ArrayElement {
                key: None,
                value: walk_expr(p),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn walk_hash_literal(pair: Pair<Rule>) -> Expression {
    let mut props = Vec::new();
    for entry in pair.into_inner() {
        if entry.as_rule() != Rule::hash_entry {
            continue;
        }
        let mut parts = entry.into_inner();
        let Some(key_pair) = parts.next() else { continue };
        let key = hash_key_text(key_pair);
        let value = parts.next().map(walk_expr).unwrap_or_else(Expression::null);
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&key),
            value,
        });
    }
    Expression::new(ExprKind::Object(props))
}

fn hash_key_text(pair: Pair<Rule>) -> String {
    let inner = pair.into_inner().next();
    match inner {
        Some(p) => {
            let raw = p.as_str();
            match p.as_rule() {
                Rule::quoted_string | Rule::single_quoted_string => raw
                    .get(1..raw.len().saturating_sub(1))
                    .unwrap_or("")
                    .to_string(),
                _ => raw.to_string(),
            }
        }
        None => String::new(),
    }
}

/// `$( … )` — a statement list whose value is the last expression.
fn walk_sub_expr(pair: Pair<Rule>) -> Expression {
    let stmts = collect_statements(pair);
    last_expression_of(stmts)
}

/// A script block is a lambda. PowerShell binds the current pipeline item to
/// `$_`, so unless the block declares its own `param(…)` the walker gives it a
/// single implicit `_` parameter — that is what makes `{ $_ -gt 2 }` receive the
/// element the shared HOF dispatch passes in.
fn walk_script_block_expr(pair: Pair<Rule>) -> Expression {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        if child.as_rule() == Rule::param_stmt {
            parse_param_stmt(child, &mut params);
            continue;
        }
        if let Ok(Some(stmt)) = parse_statement(child) {
            body.push(stmt);
        }
    }

    if params.is_empty() {
        params.push(Param {
            name: "_".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: false,
        });
    }

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(implicit_return(body)),
        is_async: false,
        captures: Vec::new(),
    })
}

/// PowerShell yields the value of a trailing expression rather than requiring
/// `return`. Rewriting that last statement here keeps the shared compiler
/// generic — the same normalization Ruby's walker does.
fn implicit_return(mut body: Vec<Statement>) -> Vec<Statement> {
    let Some(last) = body.pop() else {
        return body;
    };
    let replaced = match last.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Return(Some(expr))),
        _ => last,
    };
    body.push(replaced);
    body
}

fn collect_statements(pair: Pair<Rule>) -> Vec<Statement> {
    let mut out = Vec::new();
    for child in pair.into_inner() {
        if let Ok(Some(stmt)) = parse_statement(child) {
            out.push(stmt);
        }
    }
    out
}

fn last_expression_of(stmts: Vec<Statement>) -> Expression {
    match stmts.into_iter().last() {
        Some(stmt) => match stmt.kind {
            StmtKind::Expr(expr) => expr,
            _ => Expression::null(),
        },
        None => Expression::null(),
    }
}

fn walk_number(raw: &str) -> Expression {
    let text = raw.trim();
    let lower = text.to_lowercase();

    if let Some(hex) = lower.strip_prefix("0x") {
        if let Ok(v) = i64::from_str_radix(hex, 16) {
            return Expression::int(v);
        }
    }

    // Size suffixes: 1kb, 2mb, …
    for (suffix, factor) in [
        ("kb", 1024_i64),
        ("mb", 1024 * 1024),
        ("gb", 1024 * 1024 * 1024),
        ("tb", 1024_i64 * 1024 * 1024 * 1024),
        ("pb", 1024_i64 * 1024 * 1024 * 1024 * 1024),
    ] {
        if let Some(head) = lower.strip_suffix(suffix) {
            if let Ok(v) = head.parse::<f64>() {
                return Expression::int((v * factor as f64) as i64);
            }
        }
    }

    let stripped = lower.trim_end_matches(['l', 'd']);
    if let Ok(v) = stripped.parse::<i64>() {
        return Expression::int(v);
    }
    if let Ok(v) = stripped.parse::<f64>() {
        return Expression::float(v);
    }
    Expression::int(0)
}

fn walk_var_ref(raw: &str) -> Expression {
    let name = raw.trim_start_matches('$').trim_matches(|c| c == '{' || c == '}');
    match name.to_lowercase().as_str() {
        "true" => Expression::bool(true),
        "false" => Expression::bool(false),
        "null" => Expression::null(),
        "this" => Expression::new(ExprKind::This),
        _ => Expression::ident(scope_qualified_name(name)),
    }
}

/// PowerShell scope modifiers (`$script:x`, `$global:x`, `$local:x`, `$env:x`)
/// are not part of the variable's identity — `$script:count` and `$count` name
/// the same storage at script scope. Drop the modifier so both spellings
/// resolve to one binding; keep `env:` since it names a different store.
fn scope_qualified_name(name: &str) -> &str {
    match name.split_once(':') {
        Some((scope, rest))
            if matches!(
                scope.to_lowercase().as_str(),
                "script" | "global" | "local" | "private" | "using" | "variable"
            ) && !rest.is_empty() =>
        {
            rest
        }
        _ => name,
    }
}

fn walk_bare_word(raw: &str) -> Expression {
    match raw.to_lowercase().as_str() {
        "true" => Expression::bool(true),
        "false" => Expression::bool(false),
        "null" => Expression::null(),
        _ => Expression::ident(raw),
    }
}

/// The class name behind a `[Type]` literal, once the walker has turned it into
/// a string. The FULL dotted name is kept: `[System.Text.StringBuilder]` has to
/// stay whole so the shared namespace resolver can match it through the dotnet
/// tree-mount. Truncating to the last segment would discard exactly what
/// `resolve_profile_namespace_chain` matches on.
fn type_name_of(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) if !name.trim().is_empty() => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", type_name_of(object)?, field))
        }
        _ => None,
    }
}

/// `[System.Math]` becomes the member chain `System.Math`, the same shape the
/// C# walker produces for a dotted `System.*` name. That is what lets the shared
/// namespace resolver reach it through the dotnet tree-mount — a plain string
/// object could never resolve.
fn type_literal_expr(name: &str) -> Expression {
    let name = name.trim();
    if name.is_empty() {
        return Expression::null();
    }
    let mut segments = name.split('.').filter(|s| !s.is_empty());
    let Some(first) = segments.next() else {
        return Expression::null();
    };
    let mut expr = Expression::ident(first);
    for seg in segments {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: seg.to_string(),
            null_safe: false,
        });
    }
    expr
}

/// PowerShell's `.Count` / `.Length` are the same idea as JS `.length`, so the
/// walker rewrites them the way PHP's `count($a)` becomes `a.length`. Anything
/// else keeps its spelling and resolves through the profile.
fn normalize_member_name(name: &str) -> String {
    match name.to_lowercase().as_str() {
        "count" | "length" => "length".to_string(),
        _ => name.to_string(),
    }
}

fn type_literal_name(raw: &str) -> &str {
    raw.trim().trim_start_matches('[').trim_end_matches(']')
}

/// Map a PowerShell operator spelling onto the shared `BinOp`.
fn build_binary(op_raw: &str, left: Expression, right: Expression) -> Expression {
    let op_text = op_raw.trim().to_lowercase();
    let symbolic = matches!(op_text.as_str(), "+" | "-" | "*" | "/" | "%" | "**");

    let word = if symbolic {
        op_text.clone()
    } else {
        let stripped = op_text.trim_start_matches('-');
        // `-ceq` / `-ieq` are the case-sensitive / case-insensitive forms of
        // `-eq`; they select comparison casing, not a different operator.
        match stripped.strip_prefix('c').or_else(|| stripped.strip_prefix('i')) {
            Some(rest) if is_comparison_word(rest) => rest.to_string(),
            _ => stripped.to_string(),
        }
    };

    // `-contains` / `-notcontains` take the collection on the LEFT, the inverse
    // of `-in` / `-notin`. Swapping the operands is what makes both spellings
    // reach the same shared `In` lowering.
    let (op, left, right) = match word.as_str() {
        "contains" => (BinOp::In, right, left),
        "notcontains" => (BinOp::NotIn, right, left),
        other => {
            let op = match other {
                "+" => BinOp::Add,
                "-" => BinOp::Sub,
                "*" => BinOp::Mul,
                "/" => BinOp::Div,
                "%" => BinOp::Mod,
                "**" => BinOp::Pow,
                "eq" => BinOp::Eq,
                "ne" => BinOp::NotEq,
                "gt" => BinOp::Gt,
                "ge" => BinOp::GtEq,
                "lt" => BinOp::Lt,
                "le" => BinOp::LtEq,
                "and" => BinOp::And,
                "or" => BinOp::Or,
                "xor" => BinOp::Xor,
                "in" => BinOp::In,
                "notin" => BinOp::NotIn,
                "is" => BinOp::Is,
                "isnot" => BinOp::IsNot,
                "like" | "match" => BinOp::Like,
                "band" => BinOp::BitAnd,
                "bor" => BinOp::BitOr,
                "bxor" => BinOp::BitXor,
                "shl" => BinOp::Shl,
                "shr" => BinOp::Shr,
                "join" => {
                    return method_call_expr(left, "join", vec![right]);
                }
                "split" => {
                    return method_call_expr(left, "split", vec![right]);
                }
                "replace" => {
                    return method_call_expr(left, "replace", vec![right]);
                }
                "notlike" | "notmatch" => {
                    return Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Like,
                            left: Box::new(left),
                            right: Box::new(right),
                        })),
                    });
                }
                _ => BinOp::Eq,
            };
            (op, left, right)
        }
    };

    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn is_comparison_word(word: &str) -> bool {
    matches!(
        word,
        "eq" | "ne"
            | "gt"
            | "ge"
            | "lt"
            | "le"
            | "like"
            | "notlike"
            | "match"
            | "notmatch"
            | "contains"
            | "notcontains"
            | "in"
            | "notin"
            | "replace"
            | "split"
            | "join"
    )
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

/// The body of a `$( … )` / `${ … }` interpolation. PowerShell allows a whole
/// statement list here, so fall back to parsing statements and taking the last
/// expression when it is not a single expression.
fn parse_script_expression(raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }

    if let Some(expr) = parse_expr_fragment(text) {
        return expr;
    }

    match super::PowerShellParser::parse(Rule::program, text) {
        Ok(mut pairs) => match pairs.next() {
            Some(root) => last_expression_of(collect_statements(root)),
            None => Expression::null(),
        },
        Err(_) => Expression::string(text),
    }
}

fn is_numeric_like(raw: &str) -> bool {
    let text = raw.trim();
    if text.is_empty() {
        return false;
    }

    text.parse::<i64>().is_ok() || text.parse::<f64>().is_ok() || text == "$null"
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
