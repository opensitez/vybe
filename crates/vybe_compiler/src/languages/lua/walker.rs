use super::{LuaParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

// Helper function to convert pest span to our Span
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

// Main parse function
pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = LuaParser::parse(Rule::chunk, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    
    let mut body = Vec::new();
    let imports = Vec::new(); // Lua doesn't have imports like JS
    
    for pair in pairs {
        match pair.as_rule() {
            Rule::chunk => {
                for inner in pair.into_inner() {
                    match inner.as_rule() {
                        Rule::block => {
                            for stmt in inner.into_inner() {
                                if stmt.as_rule() == Rule::statement {
                                    body.push(walk_statement(stmt)?);
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
    
    Ok(Module {
        name: "main".to_string(),
        language: Lang::Lua,
        body,
        imports,
    })
}

// Walk a statement
fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::local_var => walk_local_var(pair)?,
        Rule::function_decl => walk_function_decl(pair)?,
        Rule::local_function => walk_local_function(pair)?,
        Rule::if_statement => walk_if_statement(pair)?,
        Rule::while_statement => walk_while_statement(pair)?,
        Rule::repeat_statement => walk_repeat_statement(pair)?,
        Rule::for_statement => walk_for_statement(pair)?,
        Rule::for_in_statement => walk_for_in_statement(pair)?,
        Rule::return_statement => walk_return_statement(pair)?,
        Rule::break_statement => walk_break_statement(pair)?,
        Rule::do_statement => walk_do_statement(pair)?,
        Rule::assignment => walk_assignment(pair)?,
        Rule::functioncall => {
            // Function call as a statement (expression statement)
            let expr = walk_expression(pair)?;
            StmtKind::Expr(expr)
        }
        _ => return Err(format!("Unhandled statement rule: {:?}", pair.as_rule())),
    };
    
    Ok(Statement::with_span(kind, span))
}

// Walk local variable declaration
fn walk_local_var(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "local" keyword
    inner.next();
    
    let mut declarations = Vec::new();
    let mut values = Vec::new();
    
    // Collect variable names
    while let Some(pair) = inner.next() {
        match pair.as_rule() {
            Rule::name => {
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(pair.as_str().to_string()),
                    type_hint: None,
                    array_bounds: None,
                    with_events: None,
                });
            }
            Rule::EQUAL => {
                // Start collecting values
                while let Some(value_pair) = inner.next() {
                    if value_pair.as_rule() == Rule::expr {
                        values.push(walk_expression(value_pair)?);
                    }
                }
                break;
            }
            _ => {}
        }
    }
    
    // Assign values to declarations
    for (i, decl) in declarations.iter_mut().enumerate() {
        if i < values.len() {
            // TODO: Handle initialization
        }
    }
    
    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Var, // Lua local is like JS var
    })
}

// Walk function declaration
fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "function" keyword
    inner.next();
    
    let name_pair = inner.next().ok_or("Missing function name")?;
    let name = name_pair.as_str().to_string();
    
    // Skip "("
    inner.next();
    
    let mut params = Vec::new();
    // Parse parameters
    while let Some(param_pair) = inner.next() {
        match param_pair.as_rule() {
            Rule::name => {
                params.push(Param {
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                });
            }
            Rule::rparen => break,
            _ => {}
        }
    }
    
    // Parse body
    let mut body = Vec::new();
    while let Some(stmt_pair) = inner.next() {
        if stmt_pair.as_rule() == Rule::block {
            for stmt in stmt_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    body.push(walk_statement(stmt)?);
                }
            }
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

// Walk if statement
fn walk_if_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "if"
    inner.next();
    
    let cond = walk_expression(inner.next().ok_or("Missing if condition")?)?;
    
    // Skip "then"
    inner.next();
    
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;
    
    // Parse then block
    while let Some(stmt_pair) = inner.next() {
        if stmt_pair.as_rule() == Rule::block {
            for stmt in stmt_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    then_body.push(walk_statement(stmt)?);
                }
            }
        } else if stmt_pair.as_rule() == Rule::"elseif" {
            let cond = walk_expression(inner.next().ok_or("Missing elseif condition")?)?;
            // Skip "then"
            inner.next();
            
            let mut elif_body = Vec::new();
            if let Some(block_pair) = inner.next() {
                if block_pair.as_rule() == Rule::block {
                    for stmt in block_pair.into_inner() {
                        if stmt.as_rule() == Rule::statement {
                            elif_body.push(walk_statement(stmt)?);
                        }
                    }
                }
            }
            
            elifs.push((cond, elif_body));
        } else if stmt_pair.as_rule() == Rule::"else" {
            if let Some(block_pair) = inner.next() {
                if block_pair.as_rule() == Rule::block {
                    let mut body = Vec::new();
                    for stmt in block_pair.into_inner() {
                        if stmt.as_rule() == Rule::statement {
                            body.push(walk_statement(stmt)?);
                        }
                    }
                    else_body = Some(body);
                }
            }
        }
    }
    
    Ok(StmtKind::If {
        cond: Box::new(cond),
        then_body,
        elifs,
        else_body,
    })
}

// Walk while statement
fn walk_while_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "while"
    inner.next();
    
    let cond = walk_expression(inner.next().ok_or("Missing while condition")?)?;
    
    // Skip "do"
    inner.next();
    
    let mut body = Vec::new();
    if let Some(block_pair) = inner.next() {
        if block_pair.as_rule() == Rule::block {
            for stmt in block_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    body.push(walk_statement(stmt)?);
                }
            }
        }
    }
    
    Ok(StmtKind::While {
        cond: Box::new(cond),
        body,
        else_body: None,
    })
}

// Walk repeat statement (repeat-until loop)
fn walk_repeat_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "repeat"
    inner.next();
    
    let mut body = Vec::new();
    if let Some(block_pair) = inner.next() {
        if block_pair.as_rule() == Rule::block {
            for stmt in block_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    body.push(walk_statement(stmt)?);
                }
            }
        }
    }
    
    // Skip "until"
    inner.next();
    
    let cond = walk_expression(inner.next().ok_or("Missing until condition")?)?;
    
    Ok(StmtKind::DoWhile {
        body,
        cond: Box::new(cond),
        until: true, // Lua repeat-until loops test at the end
    })
}

// Walk for statement (numeric for)
fn walk_for_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // For now, implement as a simple while loop
    // TODO: Implement proper numeric for semantics
    walk_while_statement(pair)
}

// Walk for-in statement (generic for)
fn walk_for_in_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "for"
    inner.next();
    
    let var_name = inner.next().ok_or("Missing for variable name")?.as_str().to_string();
    
    // Skip "in"
    inner.next();
    
    let iter = walk_expression(inner.next().ok_or("Missing iterator expression")?)?;
    
    // Skip "do"
    inner.next();
    
    let mut body = Vec::new();
    if let Some(block_pair) = inner.next() {
        if block_pair.as_rule() == Rule::block {
            for stmt in block_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    body.push(walk_statement(stmt)?);
                }
            }
        }
    }
    
    Ok(StmtKind::ForIn {
        var: var_name,
        key: None,
        iter: Box::new(iter),
        body,
        of: true, // Lua's for-in iterates values
        else_body: None,
        is_async: false,
    })
}

// Walk return statement
fn walk_return_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "return"
    inner.next();
    
    let mut values = Vec::new();
    while let Some(expr_pair) = inner.next() {
        if expr_pair.as_rule() == Rule::expr {
            values.push(walk_expression(expr_pair)?);
        }
    }
    
    Ok(StmtKind::Return(if !values.is_empty() { Some(Box::new(values[0].clone())) } else { None }))
}

// Walk break statement
fn walk_break_statement(_pair: Pair<Rule>) -> Result<StmtKind, String> {
    Ok(StmtKind::Break(BreakTarget::Implicit))
}

// Walk do statement
fn walk_do_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    
    // Skip "do"
    inner.next();
    
    let mut body = Vec::new();
    if let Some(block_pair) = inner.next() {
        if block_pair.as_rule() == Rule::block {
            for stmt in block_pair.into_inner() {
                if stmt.as_rule() == Rule::statement {
                    body.push(walk_statement(stmt)?);
                }
            }
        }
    }
    
    Ok(StmtKind::Block(body))
}

// Walk assignment statement
fn walk_assignment(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Simplified assignment - just treat as expression statement for now
    let expr = walk_expression(pair)?;
    Ok(StmtKind::Expr(expr))
}

// Walk local function (same as function_decl but local)
fn walk_local_function(pair: Pair<Rule>) -> Result<StmtKind, String> {
    walk_function_decl(pair)
}

// Walk expression
fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::expr => {
            let inner = pair.into_inner().next().ok_or("Empty expression")?;
            return walk_expression(inner);
        }
        Rule::exp => walk_binary_expr(pair)?,
        Rule::term => walk_term(pair)?,
        Rule::primary => walk_primary(pair)?,
        Rule::functioncall => walk_function_call(pair)?,
        Rule::table_constructor => walk_table_constructor(pair)?,
        Rule::var => walk_var(pair)?,
        _ => return Err(format!("Unhandled expression rule: {:?}", pair.as_rule())),
    };
    
    Ok(Expression::with_span(kind, span))
}

// Walk binary expression
fn walk_binary_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    
    let mut current = walk_term(inner.next().ok_or("Missing term in expression")?)?;
    
    while let Some(op_pair) = inner.next() {
        let op_str = op_pair.as_str();
        let op = match op_str {
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
            "&" => BinOp::BitAnd,
            "|" => BinOp::BitOr,
            "<<" => BinOp::Shl,
            ">>" => BinOp::Shr,
            _ => return Err(format!("Unknown binary operator: {}", op_str)),
        };
        
        let right = walk_term(inner.next().ok_or("Missing right operand")?)?;
        
        current = ExprKind::Binary {
            op,
            left: Box::new(Expression::with_span(current, to_span(&op_pair))),
            right: Box::new(right),
        };
    }
    
    Ok(current)
}

// Walk term (unary expression or primary)
fn walk_term(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    
    if let Some(unop_pair) = inner.next() {
        if unop_pair.as_rule() == Rule::unop {
            let op_str = unop_pair.as_str();
            let op = match op_str {
                "not" => UnaryOp::Not,
                "#" => UnaryOp::Neg, // TODO: Lua length operator
                "-" => UnaryOp::Neg,
                "~" => UnaryOp::BitNot,
                _ => return Err(format!("Unknown unary operator: {}", op_str)),
            };
            
            let expr = walk_expression(inner.next().ok_or("Missing unary operand")?)?;
            
            return Ok(Expression::with_span(
                ExprKind::Unary {
                    op,
                    expr: Box::new(expr),
                },
                to_span(&unop_pair),
            ));
        } else {
            // It's a postfix expression
            return walk_expression(unop_pair);
        }
    }
    
    Err("Empty term".to_string())
}

// Walk primary expression
fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner = pair.into_inner().next().ok_or("Empty primary")?;
    
    match inner.as_rule() {
        Rule::name => Ok(ExprKind::Ident(inner.as_str().to_string())),
        Rule::literal => walk_literal(inner),
        Rule::table_constructor => walk_table_constructor(inner),
        Rule::"(" => {
            let expr = inner.into_inner().next().ok_or("Empty parentheses")?;
            walk_expression(expr).map(|e| e.kind)
        }
        _ => Err(format!("Unhandled primary rule: {:?}", inner.as_rule())),
    }
}

// Walk literal
fn walk_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner = pair.into_inner().next().ok_or("Empty literal")?;
    
    match inner.as_rule() {
        Rule::nil_literal => Ok(ExprKind::Lit(Literal::Null)),
        Rule::boolean => {
            let value = inner.as_str() == "true";
            Ok(ExprKind::Lit(Literal::Bool(value)))
        }
        Rule::number => {
            let num_str = inner.as_str();
            if num_str.contains('.') || num_str.contains('e') || num_str.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(num_str.parse().unwrap_or(0.0))))
            } else {
                Ok(ExprKind::Lit(Literal::Int(num_str.parse().unwrap_or(0))))
            }
        }
        Rule::string => {
            let content = inner.as_str();
            // Remove quotes
            let content = if content.starts_with('"') && content.ends_with('"') {
                &content[1..content.len()-1]
            } else {
                content
            };
            Ok(ExprKind::Lit(Literal::String(content.to_string())))
        }
        _ => Err(format!("Unhandled literal rule: {:?}", inner.as_rule())),
    }
}

// Walk table constructor
fn walk_table_constructor(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // For now, return a simple object literal
    // TODO: Implement proper table constructor
    Ok(ExprKind::Lit(Literal::Object(Vec::new())))
}

// Walk var (variable access)
fn walk_var(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner = pair.into_inner().next().ok_or("Empty var")?;
    
    match inner.as_rule() {
        Rule::name => Ok(ExprKind::Ident(inner.as_str().to_string())),
        Rule::member_access => walk_member_access(inner),
        Rule::index_access => walk_index_access(inner),
        _ => Err(format!("Unhandled var rule: {:?}", inner.as_rule())),
    }
}

// Walk member access
fn walk_member_access(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    
    let object = walk_expression(inner.next().ok_or("Missing object in member access")?)?;
    // Skip "."
    inner.next();
    let property = inner.next().ok_or("Missing property name")?.as_str().to_string();
    
    Ok(ExprKind::Member {
        object: Box::new(object),
        field: property,
        null_safe: false,
    })
}

// Walk index access
fn walk_index_access(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    
    let object = walk_expression(inner.next().ok_or("Missing object in index access")?)?;
    // Skip "["
    inner.next();
    let index = walk_expression(inner.next().ok_or("Missing index in index access")?)?;
    
    Ok(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

// Walk function call
fn walk_function_call(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    
    let callee = walk_expression(inner.next().ok_or("Missing callee in function call")?)?;
    
    let mut args = Vec::new();
    if let Some(args_pair) = inner.next() {
        if args_pair.as_rule() == Rule::args {
            for arg_pair in args_pair.into_inner() {
                if arg_pair.as_rule() == Rule::expr {
                    let expr = walk_expression(arg_pair)?;
                    args.push(Argument {
                        kind: ArgKind::Positional(expr),
                        spread: false,
                    });
                }
            }
        }
    }
    
    Ok(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}