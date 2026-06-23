use super::{LuaParser, Rule};
use crate::ast::*;
use pest::iterators::Pair;
use pest::Parser;

fn to_span(pair: &Pair<Rule>) -> Span {
    let (start_line, start_col) = pair.as_span().start_pos().line_col();
    let (end_line, end_col) = pair.as_span().end_pos().line_col();
    Span {
        start_line: start_line as u32,
        start_col: start_col as u32,
        end_line: end_line as u32,
        end_col: end_col as u32,
    }
}

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = LuaParser::parse(Rule::chunk, source).map_err(|e| format!("Parse error: {e}"))?;
    let mut body = Vec::new();
    for pair in pairs {
        if pair.as_rule() == Rule::chunk {
            for inner in pair.into_inner() {
                if inner.as_rule() == Rule::block {
                    for stmt in inner.into_inner() {
                        body.push(walk_statement(stmt)?);
                    }
                }
            }
        }
    }
    let mut module = Module {
        name: "main".to_string(),
        language: Lang::Lua,
        body,
        imports: Vec::new(),
    };
    super::normalize::normalize_module(&mut module);
    Ok(module)
}

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    // Pest may still emit a `statement` wrapper depending on call site.
    let pair = if pair.as_rule() == Rule::statement {
        pair.into_inner()
            .next()
            .ok_or("empty statement wrapper")?
    } else {
        pair
    };
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::local_var => walk_local_var(pair)?,
        Rule::function_decl => walk_function_decl(pair)?,
        Rule::if_statement => walk_if_statement(pair)?,
        Rule::while_statement => walk_while_statement(pair)?,
        Rule::repeat_statement => walk_repeat_statement(pair)?,
        Rule::return_statement => walk_return_statement(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::assign_stmt => walk_assign_stmt(pair)?,
        Rule::expr => StmtKind::Expr(walk_expression(pair)?),
        other => return Err(format!("Unhandled statement rule: {other:?}")),
    };
    Ok(Statement::with_span(kind, span))
}

fn walk_local_var(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    inner.next(); // local
    let mut declarations = Vec::new();
    let mut values = Vec::new();
    let mut in_values = false;
    for p in inner {
        match p.as_rule() {
            Rule::name => {
                if in_values {
                    return Err("unexpected name in local initializer".into());
                }
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(p.as_str().to_string()),
                    type_hint: None,
                    init: None,
                    array_bounds: None,
                    with_events: false,
                });
            }
            Rule::assign => in_values = true,
            Rule::expr => values.push(walk_expression(p)?),
            _ => {}
        }
    }
    for (i, decl) in declarations.iter_mut().enumerate() {
        if let Some(val) = values.get(i) {
            decl.init = Some(val.clone());
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    inner.next(); // function
    let name = inner
        .next()
        .ok_or("missing function name")?
        .as_str()
        .to_string();
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::name => {
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
            Rule::func_body | Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            _ => {}
        }
    }
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        body,
        modifiers: Modifiers::default(),
        is_async: false,
        is_generator: false,
        is_sub: false,
        handles: Vec::new(),
        return_type: None,
    })
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        let stmt_pair = match p.as_rule() {
            Rule::stmt_not_else | Rule::stmt_not_end => p
                .into_inner()
                .next()
                .ok_or("empty guarded statement")?,
            Rule::statement => p,
            other => return Err(format!("unexpected block item: {other:?}")),
        };
        body.push(walk_statement(stmt_pair)?);
    }
    Ok(body)
}

fn walk_if_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;
    let mut pending_elif_cond = None;
    for p in inner {
        match p.as_rule() {
            Rule::expr => {
                let e = walk_expression(p)?;
                if cond.is_none() {
                    cond = Some(e);
                } else {
                    pending_elif_cond = Some(e);
                }
            }
            Rule::then_block | Rule::else_block | Rule::block => {
                let stmts = walk_block(p)?;
                if then_body.is_empty() && pending_elif_cond.is_none() {
                    then_body = stmts;
                } else if let Some(elif_cond) = pending_elif_cond.take() {
                    elifs.push((elif_cond, stmts));
                } else {
                    else_body = Some(stmts);
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::If {
        cond: cond.ok_or("missing if condition")?,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_while_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let mut cond = None;
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::expr => cond = Some(walk_expression(p)?),
            Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            _ => {}
        }
    }
    Ok(StmtKind::While {
        cond: cond.ok_or("missing while condition")?,
        body,
        else_body: None,
    })
}

fn walk_repeat_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let mut body = Vec::new();
    let mut cond = None;
    for p in inner {
        match p.as_rule() {
            Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            Rule::expr => cond = Some(walk_expression(p)?),
            _ => {}
        }
    }
    Ok(StmtKind::DoWhile {
        body,
        cond: cond.ok_or("missing until condition")?,
        until: true,
    })
}

fn walk_return_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::expr {
            values.push(walk_expression(p)?);
        }
    }
    Ok(StmtKind::Return(values.into_iter().next()))
}

fn walk_assign_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let target = walk_expression(inner.next().ok_or("missing assignment target")?)?;
    inner.next(); // =
    let value = walk_expression(inner.next().ok_or("missing assignment value")?)?;
    Ok(StmtKind::Assign {
        targets: vec![target],
        value,
    })
}

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::call_expression => return walk_call_expression(pair),
        Rule::primary => return walk_primary(pair),
        Rule::postfix => {
            let inner = pair
                .into_inner()
                .next()
                .ok_or("empty postfix")?;
            return walk_call_expression(inner);
        }
        _ => {}
    }

    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expression(inner.remove(0)).map(|e| e.kind);
    }
    match rule {
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::compare_expr => walk_compare_chain(inner),
        Rule::concat_expr => walk_binary_chain(inner, |_| BinOp::Concat),
        Rule::additive | Rule::multiplicative | Rule::pow_expr => walk_binary_chain_with_ops(inner),
        Rule::unary => walk_unary_from_inner(inner),
        Rule::expr => walk_expression(inner.remove(0)).map(|e| e.kind),
        other => Err(format!("Unhandled expression rule: {other:?}")),
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_lua_expr_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            left = Expression::new(ExprKind::Binary {
                op: op_fn(""),
                left: Box::new(left),
                right: Box::new(right),
            });
        }
    }
    Ok(left.kind)
}

fn walk_compare_chain(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        if items[i].as_rule() == Rule::compare_op {
            let op = parse_binop(items[i].as_str().trim())?;
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
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
        if is_lua_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim())?;
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            }
        } else if is_lua_expr_rule(p.as_rule()) {
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(left),
                right: Box::new(right),
            });
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn is_lua_expr_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::or_expr
            | Rule::and_expr
            | Rule::compare_expr
            | Rule::concat_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::pow_expr
            | Rule::unary
            | Rule::postfix
            | Rule::call_expression
            | Rule::primary
            | Rule::expr
    )
}

fn is_lua_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::mul_op
            | Rule::pow_op
            | Rule::compare_op
            | Rule::PLUS
            | Rule::MINUS
            | Rule::STAR
            | Rule::SLASH
            | Rule::DOUBLESLASH
            | Rule::PERCENT
            | Rule::CARET
            | Rule::CONCAT
            | Rule::EQ
            | Rule::NE
            | Rule::LT
            | Rule::GT
            | Rule::LE
            | Rule::GE
    )
}

fn walk_unary_from_inner(mut inner: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let first = inner.remove(0);
    if first.as_rule() == Rule::unop {
        let op_str = first.as_str();
        let operand = walk_expression(inner.remove(0))?;
        if op_str == "#" {
            return Ok(ExprKind::Member {
                object: Box::new(operand),
                field: "length".to_string(),
                null_safe: false,
            });
        }
        let op = parse_unop(op_str)?;
        return Ok(ExprKind::Unary {
            op,
            expr: Box::new(operand),
        });
    }
    walk_expression(first).map(|e| e.kind)
}

/// Walk `primary ~ call_chain*` — same shape as JS `walk_call_chain`.
fn walk_call_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty call_expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain {
            continue;
        }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with('(') {
            let mut args = Vec::new();
            for arg in chain_inner {
                if arg.as_rule() == Rule::expr {
                    args.push(Argument::positional(walk_expression(arg)?));
                }
            }
            expr = Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args,
                optional: false,
            });
        } else if chain_src.starts_with('.') {
            let field = chain_inner
                .iter()
                .find(|p| p.as_rule() == Rule::name)
                .ok_or("missing member name")?
                .as_str()
                .to_string();
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field,
                null_safe: false,
            });
        } else if chain_src.starts_with('[') {
            let index = chain_inner
                .iter()
                .find(|p| p.as_rule() == Rule::expr)
                .map(|p| walk_expression(p.clone()))
                .transpose()?
                .ok_or("missing index")?;
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(index),
                null_safe: false,
            });
        }
    }
    Ok(expr.kind)
}

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty primary")?;
    match first.as_rule() {
        Rule::KW_TRUE => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::KW_FALSE => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::KW_NIL => Ok(ExprKind::Lit(Literal::Null)),
        Rule::number => {
            let v: f64 = first.as_str().parse().unwrap_or(0.0);
            Ok(ExprKind::Lit(Literal::Float(v)))
        }
        Rule::string => {
            let raw = first.as_str();
            let content = if raw.len() >= 2 && raw.starts_with('"') && raw.ends_with('"') {
                &raw[1..raw.len() - 1]
            } else {
                raw
            };
            Ok(ExprKind::Lit(Literal::Str(content.to_string())))
        }
        Rule::name => Ok(ExprKind::Ident(first.as_str().to_string())),
        Rule::lparen => {
            let expr = inner
                .find(|p| p.as_rule() == Rule::expr)
                .ok_or("empty parentheses")?;
            walk_expression(expr).map(|e| e.kind)
        }
        other => Err(format!("Unhandled primary: {other:?}")),
    }
}

fn parse_binop(s: &str) -> Result<BinOp, String> {
    Ok(match s {
        "and" => BinOp::And,
        "or" => BinOp::Or,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "~=" => BinOp::NotEq,
        "==" => BinOp::Eq,
        ".." => BinOp::Concat,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "//" => BinOp::FloorDiv,
        "%" => BinOp::Mod,
        "^" => BinOp::Pow,
        _ => return Err(format!("unknown binop: {s}")),
    })
}

fn parse_unop(s: &str) -> Result<UnaryOp, String> {
    Ok(match s {
        "not" => UnaryOp::Not,
        "-" => UnaryOp::Neg,
        "~" => UnaryOp::BitNot,
        _ => return Err(format!("unknown unop: {s}")),
    })
}
