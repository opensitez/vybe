use pest::Parser;
use pest::iterators::Pair;

use super::{Rule, VbParser};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs = VbParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::imports_statement => imports.push(parse_imports_statement(inner)?),
                Rule::option_directive => {}
                Rule::statement_line => {
                    for stmt_pair in inner.into_inner() {
                        if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                            continue;
                        }
                        if stmt_pair.as_rule() == Rule::module_decl {
                            body.push(parse_module_decl(stmt_pair)?);
                        } else if stmt_pair.as_rule() == Rule::namespace_decl {
                            body.push(parse_namespace_decl(stmt_pair)?);
                        } else if stmt_pair.as_rule() == Rule::dim_statement {
                            body.push(parse_statement(stmt_pair)?);
                        } else if let Some(decl_stmt) = try_parse_declaration(stmt_pair.clone())? {
                            body.push(decl_stmt);
                        } else {
                            body.push(parse_statement(stmt_pair)?);
                        }
                    }
                }
                Rule::NEWLINE | Rule::EOI => {}
                _ => {}
            }
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::VB,
        body,
        imports,
    })
}

pub fn parse_expression_str(source: &str) -> Result<Expression, String> {
    let mut pairs = VbParser::parse(Rule::expression, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let pair = pairs.next().ok_or_else(|| "Missing VB expression".to_string())?;
    parse_expression(pair)
}

fn normalize_vb_identifier(name: &str) -> String {
    name.strip_prefix('[')
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(name)
        .to_string()
}

fn canonicalize_call(_name: &str, _arguments: &[Argument]) -> Option<Expression> {
    None
}

fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    let is_class_static = matches!(
        &object.kind,
        ExprKind::Ident(n) if n.chars().next().is_some_and(|c| c.is_ascii_uppercase())
    );

    if matches!(name, "Keys" | "Values") && !is_class_static {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(object),
                field: name.to_string(),
                null_safe: false,
            })),
            args: vec![],
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

fn emit_vb_object_init_iife(new_call: Expression, props: Vec<(String, Expression)>) -> Expression {
    let type_hint = match &new_call.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        },
        _ => None,
    };
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint,
                init: Some(new_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (name, value) in props {
        let assign = Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__obj")),
                field: name,
                null_safe: false,
            })),
            value: Box::new(value),
        });
        body.push(Statement::with_span(StmtKind::Expr(assign), Span::default()));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__obj"))),
        Span::default(),
    ));

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![],
        optional: false,
    })
}

fn try_parse_declaration(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let stmt = match pair.as_rule() {
        Rule::sub_decl => Some(parse_sub_decl(pair)?),
        Rule::function_decl => Some(parse_function_decl(pair)?),
        Rule::module_decl => Some(parse_module_decl(pair)?),
        Rule::namespace_decl => Some(parse_namespace_decl(pair)?),
        Rule::class_decl => Some(parse_class_decl(pair)?),
        Rule::interface_decl => Some(parse_interface_decl(pair)?),
        Rule::structure_decl => Some(parse_structure_decl(pair)?),
        Rule::enum_decl => Some(parse_enum_decl(pair)?),
        _ => None,
    };
    Ok(stmt)
}

/*

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs = VbParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::imports_statement => imports.push(parse_imports_statement(inner)?),
                Rule::statement_line => {
                    for stmt_pair in inner.into_inner() {
                        if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                            continue;
                            let elements: Vec<Expression> = p.into_inner()
                                .filter(|e| e.as_rule() == Rule::expression)
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            let mut all_args = args;
                            for elem in elements {
                                all_args.push(Argument::positional(elem));
                            }
                            return Ok(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args: all_args,
                            }, span));
                        }
                        Rule::with_initializer => {
                            let mut members = Vec::new();
                            for mi in p.into_inner() {
                                if mi.as_rule() != Rule::member_initializer {
                                    continue;
                                }
                                let mut mi_inner = mi.into_inner();
                                let prop_name = mi_inner.next().unwrap().as_str().to_string();
                                let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                                members.push((prop_name, prop_expr));
                            }
                            return Ok(emit_vb_object_init_iife(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args,
                            }, span), members));
                        }
                        _ => {}
                    }
                }
                if let Some(elements) = array_init {
                    ExprKind::Array(
                        elements.into_iter().map(|expr| ArrayElement {
                            key: None,
                            value: expr,
                            spread: false,
                            by_ref: false,
                        }).collect(),
                    )
                } else {
                    ExprKind::New {
                        class: Box::new(Expression::ident(&class_name)),
                        args,
                    }
                }
            }
            Rule::if_expression => {
                let mut inner = pair.into_inner();
                let first = parse_expression(inner.next().unwrap())?;
                let second = parse_expression(inner.next().unwrap())?;
                let third = inner.next().map(parse_expression).transpose()?;
                match third {
                    Some(else_expr) => ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    },
                    None => ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    },
                }
            }
            Rule::addressof_expr => {
                let inner = pair.into_inner();
                let mut name = String::new();
                for p in inner {
                    if p.as_rule() == Rule::dotted_identifier {
                        name = p.as_str().to_string();
                    }
                }
                ExprKind::AddressOf(name)
            }
            Rule::me_keyword => ExprKind::This,
            Rule::dot_call_statement => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments: Vec<Argument> = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("dot_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::dot_member_access => {
                let inner = pair.into_inner();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_access => {
                let mut inner = pair.into_inner();
                let _me = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::This);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::mybase_member_access => {
                let mut inner = pair.into_inner();
                let _mybase = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::Super);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = vec![];
                let mut arguments: Vec<Argument> = vec![];
                for p in inner {
                    match p.as_rule() {
                        Rule::me_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }

                if identifiers.is_empty() {
                    return Err("me_member_call needs at least one identifier".to_string());
                }

                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::This);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }

                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::mybase_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = vec![];
                let mut arguments: Vec<Argument> = vec![];
                for p in inner {
                    match p.as_rule() {
                        Rule::mybase_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }

                if identifiers.is_empty() {
                    return Err("mybase_member_call needs at least one identifier".to_string());
                }

                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Super);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }

                if identifiers.len() == 1 {
                    ExprKind::SuperCall {
                        method: Some(method_name),
                        args: arguments,
                    }
                } else {
                    let callee = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: method_name,
                        null_safe: false,
                    });
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args: arguments,
                        optional: false,
                    }
                }
            }
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };

        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "and" | "andalso" => BinOp::And,
                    "or" | "orelse" => BinOp::Or,
                    "xor" => BinOp::Xor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left),
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}

*/
fn parse_sub_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut body = Vec::new();
    let mut handles: Vec<String> = Vec::new();
    let mut is_async = false;
    let mut is_extension = false;
    let mut is_overridable = false;
    let mut is_overrides = false;
    let mut is_must_override = false;
    let mut is_shared = false;
    let mut is_not_overridable = false;

    for p in inner {
        match p.as_rule() {
            Rule::extension_attribute => is_extension = true,
            Rule::visibility_modifier => visibility = parse_visibility(p.as_str()),
            Rule::async_kw => is_async = true,
            Rule::sub_modifier_keyword => {
                let kw = p.as_str().to_lowercase();
                match kw.as_str() {
                    "overrides" => is_overrides = true,
                    "overridable" => is_overridable = true,
                    "mustoverride" => is_must_override = true,
                    "shared" => is_shared = true,
                    "notoverridable" => is_not_overridable = true,
                    _ => {}
                }
            }
            Rule::identifier | Rule::member_identifier | Rule::sub_name => name = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::sub_block_body => {
                body.extend(parse_block(p)?);
            }
            Rule::sub_inline_body => {
                for stmt_pair in p.into_inner() {
                    match stmt_pair.as_rule() {
                        Rule::sub_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => body.push(parse_statement(stmt_pair)?),
                    }
                }
            }
            Rule::handles_clause => {
                for hp in p.into_inner() {
                    if hp.as_rule() == Rule::dotted_identifier {
                        handles.push(hp.as_str().to_string());
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name,
        params: parameters,
        return_type: None,
        body,
        modifiers: Modifiers {
            visibility,
            is_static: is_shared,
            is_abstract: is_must_override,
            is_virtual: is_overridable,
            is_override: is_overrides,
            is_readonly: false,
            is_shared,
            is_extension,
            is_overloads: false,
            is_not_overridable,
            decorators: vec![],
        },
        handles,
        is_async,
        is_generator: false,
        is_sub: true,
    }, span))
}

fn parse_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut handles: Vec<String> = Vec::new();
    let mut is_async = false;
    let mut is_extension = false;
    let mut is_overridable = false;
    let mut is_overrides = false;
    let mut is_must_override = false;
    let mut is_shared = false;
    let mut is_not_overridable = false;

    for p in inner {
        match p.as_rule() {
            Rule::extension_attribute => is_extension = true,
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::async_kw => is_async = true,
            Rule::sub_modifier_keyword => {
                let kw = p.as_str().to_lowercase();
                match kw.as_str() {
                    "overrides" => is_overrides = true,
                    "overridable" => is_overridable = true,
                    "mustoverride" => is_must_override = true,
                    "shared" => is_shared = true,
                    "notoverridable" => is_not_overridable = true,
                    _ => {}
                }
            },
            Rule::identifier | Rule::member_identifier | Rule::function_name => name = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::func_block_body => {
                body.extend(parse_block(p)?);
            }
            Rule::func_inline_body => {
                for stmt_pair in p.into_inner() {
                    match stmt_pair.as_rule() {
                        Rule::func_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => body.push(parse_statement(stmt_pair)?),
                    }
                }
            }
            Rule::handles_clause => {
                for hp in p.into_inner() {
                    if hp.as_rule() == Rule::dotted_identifier {
                        handles.push(hp.as_str().to_string());
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name,
        params: parameters,
        return_type,
        body,
        modifiers: Modifiers {
            visibility,
            is_static: is_shared,
            is_abstract: is_must_override,
            is_virtual: is_overridable,
            is_override: is_overrides,
            is_readonly: false,
            is_shared,
            is_extension,
            is_overloads: false,
            is_not_overridable,
            decorators: vec![],
        },
        handles,
        is_async,
        is_generator: false,
        is_sub: false,
    }, span))
}

fn parse_module_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::sub_decl => members.push(ClassMember::Method(Box::new(parse_sub_decl(p)?))),
            Rule::function_decl => members.push(ClassMember::Method(Box::new(parse_function_decl(p)?))),
            Rule::const_statement => {
                let (vis, decl) = parse_const_statement(p)?;
                let init = decl.init.unwrap_or_else(|| Expression::null());
                let name = match decl.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Const {
                    name,
                    type_hint: decl.type_hint,
                    value: init,
                    visibility: vis,
                });
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: Modifiers::default(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let modifiers = parse_field_modifiers(&p);
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::class_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_class_decl(p)?)));
            }
            Rule::enum_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_enum_decl(p)?)));
            }
            Rule::NEWLINE | Rule::module_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::ModuleDecl {
        name,
        members,
        visibility: Visibility::Public,
    }, span))
}

fn parse_field_modifiers(pair: &Pair<Rule>) -> Modifiers {
    let mut modifiers = Modifiers::default();

    for field_part in pair.clone().into_inner() {
        match field_part.as_rule() {
            Rule::visibility_modifier => {
                modifiers.visibility = parse_visibility(field_part.as_str());
            }
            Rule::sub_modifier_keyword if field_part.as_str().eq_ignore_ascii_case("shared") => {
                modifiers.is_static = true;
                modifiers.is_shared = true;
            }
            _ => {}
        }
    }

    modifiers
}

/// Parse `Imports [alias =] dotted.path`
fn parse_imports_statement(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut alias: Option<String> = None;
    let mut path = String::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::imports_alias => {
                // imports_alias = { identifier ~ "=" }
                if let Some(id) = p.into_inner().next() {
                    alias = Some(id.as_str().to_string());
                }
            }
            Rule::dotted_identifier => {
                path = p.as_str().to_string();
            }
            Rule::NEWLINE => {}
            _ => {}
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span,
    })
}

/// Parse `Namespace dotted.name ... End Namespace`
fn parse_namespace_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::dotted_identifier => {
                name = p.as_str().to_string();
            }
            Rule::class_decl => {
                body.push(parse_class_decl(p)?);
            }
            Rule::module_decl => {
                body.push(parse_module_decl(p)?);
            }
            Rule::enum_decl => {
                body.push(parse_enum_decl(p)?);
            }
            Rule::namespace_decl => {
                // Nested namespace
                body.push(parse_namespace_decl(p)?);
            }
            Rule::interface_decl => {
                body.push(parse_interface_decl(p)?);
            }
            Rule::structure_decl => {
                body.push(parse_structure_decl(p)?);
            }
            Rule::NEWLINE | Rule::namespace_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::NamespaceDecl { name, body }, span))
}

/// Parse an auto-implemented property into a VarDeclarator (field), since it's syntactic sugar.
fn parse_auto_property_as_field(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut var_type: Option<String> = None;
    let mut initializer = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_name => var_type = Some(p.as_str().to_string()),
            Rule::expression => initializer = Some(parse_expression(p)?),
            // Skip visibility, ReadOnly, WriteOnly keywords
            _ => {}
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: var_type,
        init: initializer,
        array_bounds: None,
        with_events: false,
    })
}

fn parse_class_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut name = String::new();
    let mut is_partial = false;
    let mut visibility = Visibility::Public;
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut is_must_inherit = false;
    let mut is_not_inheritable = false;

    for p in inner {
        match p.as_rule() {
            Rule::partial_keyword => is_partial = true,
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::must_inherit_keyword => is_must_inherit = true,
            Rule::not_inheritable_keyword => is_not_inheritable = true,
            Rule::inherits_statement => {
                if let Some(type_pair) = p.into_inner().next() {
                    // Strip qualification: `System.Windows.Forms.Form` → `Form`.
                    // The canonical AST stores unqualified parent names because
                    // the compiler resolves them via globals (where dotnet
                    // classes are installed by `register_dotnet_classes`). User
                    // classes live in the same flat global namespace, so this
                    // matches both .NET BCL parents and user-defined parents.
                    let qualified = type_pair.as_str().to_string();
                    let unqualified = qualified.rsplit('.').next().unwrap_or(&qualified).to_string();
                    parents.push(unqualified);
                }
            }
            Rule::implements_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        interfaces.push(tp.as_str().to_string());
                    }
                }
            }
            Rule::identifier => name = p.as_str().to_string(),
            Rule::property_decl => {
                members.push(parse_property_decl_to_member(p)?);
            }
            Rule::auto_property_decl => {
                // Auto-implemented property → treat as a field
                let d = parse_auto_property_as_field(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers: Modifiers::default(),
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::sub_decl => {
                let sub_stmt = parse_sub_decl(p)?;
                // Check if this is a constructor (New)
                let is_ctor = match &sub_stmt.kind {
                    StmtKind::FunctionDecl { name, .. } => name == "New",
                    _ => false,
                };
                if is_ctor {
                    match sub_stmt.kind {
                        StmtKind::FunctionDecl { params, body, modifiers, .. } => {
                            members.push(ClassMember::Constructor {
                                params,
                                body,
                                base_args: None,
                                initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                                visibility: modifiers.visibility,
                            });
                        }
                        _ => unreachable!(),
                    }
                } else {
                    members.push(ClassMember::Method(Box::new(sub_stmt)));
                }
            }
            Rule::function_decl => {
                members.push(ClassMember::Method(Box::new(parse_function_decl(p)?)));
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: Modifiers::default(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let modifiers = parse_field_modifiers(&p);
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::event_decl => {
                let ev = parse_event_decl_to_member(p)?;
                members.push(ev);
            }
            Rule::class_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_class_decl(p)?)));
            }
            Rule::enum_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_enum_decl(p)?)));
            }
            Rule::NEWLINE | Rule::class_end => {}
            _ => {}
        }
    }

    // Inject canonical AddHandler statements at the END of the constructor
    // body for every class method that has a `Handles` clause. This is the
    // walker normalization that turns VB-specific `Handles ctrl.Event` into
    // the same canonical `StmtKind::AddHandler` that C# `+=` (and JS / Dart /
    // Python frontends) will produce. The compiler then has a single emit
    // path for events regardless of source language.
    inject_handles_into_constructor(&mut members);

    // Inject implicit `MyBase.New()` at the START of every constructor body
    // when this class has an `Inherits` clause and the body doesn't already
    // start with an explicit `MyBase.New(...)`. This matches real VB.NET
    // semantics: the runtime implicitly calls the parameterless parent ctor
    // before the body runs, and the VB compiler errors if no parameterless
    // parent ctor is accessible. By doing this here we keep the canonical
    // AST uniform — compile_class sees a body that always starts with the
    // base call, the same as Pascal `inherited Create(...)` and JS
    // `super(...)`. The compiler-side logic doesn't need a VB-specific
    // case.
    //
    // Also stamp `Me.__control_name = "<lowercased class name>"` immediately
    // after the base call, so any subsequent property writes (e.g.
    // `Me.Text = "X"`) mirror to the gui state under the user-meaningful
    // key. The base ctor (e.g. `Form()`) wires the underlying widget which
    // has its own auto-generated `__control_name`; this re-stamp overrides
    // it with the canonical "lowercased subclass name" form that real
    // WinForms users (and the existing Vybe form runner) expect.
    if !parents.is_empty() {
        inject_implicit_mybase_new(&mut members, &name);
    }

    Ok(Statement::with_span(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers {
            visibility,
            is_partial,
            is_abstract: is_must_inherit,
            is_sealed: is_not_inheritable,
            is_static: false,
        },
        decorators: vec![],
    }, span))
}

/// Inject `MyBase.New()` at the start of every constructor body that doesn't
/// already start with an explicit base call. Matches real VB.NET semantics
/// for `Inherits` classes.
///
/// "Already starts with one" is checked structurally — only the FIRST
/// statement is examined, because that's the only legal position for an
/// explicit `MyBase.New(...)` in VB. (VB.NET errors if `MyBase.New` appears
/// anywhere else in the body.)
///
/// If the class has no constructor at all, a default one is synthesized
/// containing just the `MyBase.New()` call.
fn inject_implicit_mybase_new(members: &mut Vec<ClassMember>, class_name: &str) {
    let lowered = class_name.to_lowercase();

    let mybase_new = || -> Statement {
        // SuperCall { method: Some("New"), args: [] } — same shape the
        // mybase_member_call walker arm produces for explicit `MyBase.New()`.
        Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::SuperCall { method: Some("New".to_string()), args: Vec::new() },
        )))
    };

    // Me.__control_name = "<lowercased class name>"
    // This is the .NET-canonical "self identity" stamp. The base ctor (e.g.
    // `Form()`) wired up the underlying widget with its own auto-generated
    // name; we override it here so user property writes mirror to gui state
    // under the subclass name the rest of the system uses.
    let stamp_control_name = || -> Statement {
        let target = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Me")),
            field: "__control_name".to_string(),
            null_safe: false,
        });
        let value = Expression::string(&lowered);
        Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
        })
    };

    let starts_with_mybase_new = |body: &[Statement]| -> bool {
        match body.first().map(|s| &s.kind) {
            Some(StmtKind::Expr(e)) => matches!(
                &e.kind,
                ExprKind::SuperCall { .. }
            ),
            _ => false,
        }
    };

    let has_ctor = members.iter().any(|m| matches!(m, ClassMember::Constructor { .. }));
    if !has_ctor {
        // Synthesize a default ctor that just calls MyBase.New() and stamps
        // the canonical control name.
        members.push(ClassMember::Constructor {
            params: Vec::new(),
            body: vec![mybase_new(), stamp_control_name()],
            base_args: None,
            initializer_target: crate::ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
        return;
    }
    for m in members.iter_mut() {
        if let ClassMember::Constructor { body, .. } = m {
            if !starts_with_mybase_new(body) {
                body.insert(0, mybase_new());
                body.insert(1, stamp_control_name());
            } else {
                // Body already starts with MyBase.New() — insert the stamp
                // immediately after it.
                body.insert(1, stamp_control_name());
            }
        }
    }
}

/// Walk the class members; for every Method with `handles: ["ctrl.Event", ...]`,
/// build a canonical `AddHandler { control, event, handler: Me.method_name }`
/// statement and append it to the constructor body. If no constructor exists,
/// inject an empty one. Strips the `handles` field from the method afterward
/// so the compiler doesn't double-process it.
fn inject_handles_into_constructor(members: &mut Vec<ClassMember>) {
    // First pass: collect (handler_method_name, handles_list) and clear them
    // off the methods so the compile_function_decl path doesn't re-emit.
    let mut to_inject: Vec<(String, Vec<String>)> = Vec::new();
    for m in members.iter_mut() {
        if let ClassMember::Method(stmt) = m {
            if let StmtKind::FunctionDecl { name: mname, handles, modifiers, .. } = &mut stmt.kind {
                if !handles.is_empty() && !modifiers.is_static {
                    to_inject.push((mname.clone(), std::mem::take(handles)));
                }
            }
        }
    }
    if to_inject.is_empty() { return; }

    // Build the AddHandler statements.
    let mut new_stmts: Vec<Statement> = Vec::new();
    for (method_name, handles) in &to_inject {
        for h in handles {
            let (control, event) = split_event_target(h);
            // The handler is `Me.<method>` — a Member access on the class self.
            let handler = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Me")),
                field: method_name.clone(),
                null_safe: false,
            });
            new_stmts.push(Statement::new(crate::common::events::add_handler_stmt(
                control,
                event,
                handler,
            )));
        }
    }

    // Find the constructor (or create one) and append the AddHandler statements.
    // VB constructors can appear either as an explicit `Sub New()` method or as a
    // dedicated `ClassMember::Constructor` node, depending on which parser path
    // produced the member. Handles normalization must attach to whichever form the
    // class already uses; otherwise `compile_class` can end up ignoring the
    // injected AddHandler body by selecting the explicit `Sub New()` body first.
    let has_explicit_new = members.iter().any(|m| {
        matches!(m,
            ClassMember::Method(stmt)
                if matches!(&stmt.kind,
                    StmtKind::FunctionDecl { name, .. } if name.eq_ignore_ascii_case("new")
                )
        )
    });
    let has_ctor = members.iter().any(|m| matches!(m, ClassMember::Constructor { .. }));
    if !has_ctor && !has_explicit_new {
        members.push(ClassMember::Constructor {
            params: Vec::new(),
            body: new_stmts,
            base_args: None,  // VB walker injects MyBase.New() into the body; no base_args needed here
            initializer_target: crate::ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    } else {
        for m in members.iter_mut() {
            match m {
                ClassMember::Constructor { body, .. } => {
                    body.extend(new_stmts.drain(..));
                    break;
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name, body, .. } = &mut stmt.kind {
                        if name.eq_ignore_ascii_case("new") {
                            body.extend(new_stmts.drain(..));
                            break;
                        }
                    }
                }
                _ => {}
            }
        }
    }
}

fn parse_property_decl_to_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut _parameters = Vec::new();
    let mut return_type: Option<String> = None;
    let mut getter = None;
    let mut setter = None;

    for p in inner {
        match p.as_str().to_lowercase().as_str() {
            "public" => visibility = Visibility::Public,
            "private" => visibility = Visibility::Private,
            _ => {
                match p.as_rule() {
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::param_list => _parameters = parse_param_list(p)?,
                    Rule::type_name => return_type = Some(p.as_str().to_string()),
                    Rule::property_get => getter = Some(parse_property_get(p)?),
                    Rule::property_set => setter = Some(parse_property_set(p)?),
                    _ => {}
                }
            }
        }
    }

    Ok(ClassMember::Property {
        name,
        type_hint: return_type,
        getter,
        setter,
        is_auto: false,
        modifiers: Modifiers {
            visibility,
            ..Modifiers::default()
        },
    })
}

fn parse_property_get(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt_pair in pair.into_inner() {
        if stmt_pair.as_rule() == Rule::statement_line {
            for s in stmt_pair.into_inner() {
                if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                    body.push(parse_statement(s)?);
                }
            }
        }
    }
    Ok(body)
}

fn parse_property_set(pair: Pair<Rule>) -> Result<PropertySetter, String> {
    let mut inner = pair.into_inner();
    let param = parse_parameter(inner.next().unwrap())?; // Set(ByVal value As Type)

    let mut body = Vec::new();
    for stmt_pair in inner {
        if stmt_pair.as_rule() == Rule::statement_line {
            for s in stmt_pair.into_inner() {
                if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                    body.push(parse_statement(s)?);
                }
            }
        }
    }

    Ok(PropertySetter { param, body })
}

fn parse_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    pair.into_inner().map(parse_parameter).collect()
}

fn parse_parameter(pair: Pair<Rule>) -> Result<Param, String> {
    let inner = pair.into_inner();
    let mut pass_by = PassBy::Value;
    let mut name = String::new();
    let mut param_type: Option<String> = None;
    let mut is_optional = false;
    let mut default_value = None;
    let mut is_nullable = false;
    let mut is_param_array = false;

    for p in inner {
        match p.as_rule() {
            Rule::pass_type_keyword => {
                let text = p.as_str().to_lowercase();
                if text == "byval" {
                    pass_by = PassBy::Value;
                } else {
                    pass_by = PassBy::Ref;
                }
            }
            Rule::optional_keyword => {
                is_optional = true;
            }
            Rule::paramarray_keyword => {
                is_param_array = true;
                pass_by = PassBy::Value; // ParamArray is always ByVal
            }
            Rule::identifier => {
                name = p.as_str().to_string();
            }
            Rule::type_name => param_type = Some(p.as_str().to_string()),
            Rule::nullable_marker => is_nullable = true,
            Rule::expression => default_value = Some(parse_expression(p)?),
            _ => {}
        }
    }

    Ok(Param {
        name,
        type_hint: param_type,
        default: default_value,
        pass_by,
        is_rest: is_param_array,
        is_kwargs: false,
        is_optional,
        is_nullable,
    })
}

fn parse_array_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let elements = pair.into_inner()
        .map(|p| parse_expression(p).map(|value| ArrayElement {
            key: None,
            value,
            spread: false,
            by_ref: false,
        }))
        .collect::<Result<Vec<_>, _>>()?;
    Ok(Expression::with_span(ExprKind::Array(elements), span))
}

fn parse_const_statement(pair: Pair<Rule>) -> Result<(Visibility, VarDeclarator), String> {
    let mut visibility = Visibility::Private;
    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::visibility_modifier => visibility = parse_visibility(p.as_str()),
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::expression => init = Some(parse_expression(p)?),
            Rule::array_literal => init = Some(parse_array_literal(p)?),
            _ => {}
        }
    }

    Ok((visibility, VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint,
        init,
        array_bounds: None,
        with_events: false,
    }))
}

fn parse_array_bounds_pair(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    match pair.as_rule() {
        Rule::array_rank_spec => pair
            .into_inner()
            .next()
            .map(parse_array_bounds_pair)
            .transpose()
            .map(|bounds| bounds.unwrap_or_default()),
        Rule::array_bounds => pair
            .into_inner()
            .map(parse_expression)
            .collect::<Result<Vec<_>, _>>(),
        _ => Err(format!("Unexpected array bounds rule: {:?}", pair.as_rule())),
    }
}

fn parse_dim_statement(pair: Pair<Rule>) -> Result<Vec<VarDeclarator>, String> {
    let mut decls = Vec::new();
    for part in pair.into_inner() {
        if part.as_rule() != Rule::dim_declaration_part {
            continue;
        }

        let mut name = String::new();
        let mut type_hint = None;
        let mut init = None;
        let mut array_bounds = None;
        let mut array_rank_count = 0usize;
        let mut ctor_args = Vec::new();
        let mut is_new = false;

        for p in part.into_inner() {
            match p.as_rule() {
                Rule::identifier => name = p.as_str().to_string(),
                Rule::array_rank_spec => {
                    array_rank_count += 1;
                    if array_bounds.is_none() {
                        array_bounds = Some(parse_array_bounds_pair(p)?);
                    }
                }
                Rule::array_bounds => {
                    array_bounds = Some(parse_array_bounds_pair(p)?);
                }
                Rule::type_name => type_hint = Some(p.as_str().to_string()),
                Rule::dim_new_keyword => is_new = true,
                Rule::argument_list => ctor_args = parse_argument_list(p)?,
                Rule::expression => init = Some(parse_expression(p)?),
                Rule::array_literal => init = Some(parse_array_literal(p)?),
                Rule::from_initializer => {
                    let mut args = ctor_args.clone();
                    for elem in p.into_inner().filter(|e| e.as_rule() == Rule::expression) {
                        args.push(Argument::positional(parse_expression(elem)?));
                    }
                    if let Some(class_name) = &type_hint {
                        init = Some(Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident(class_name)),
                            args,
                        }));
                    }
                }
                Rule::with_initializer => {
                    let mut members = Vec::new();
                    for mi in p.into_inner() {
                        if mi.as_rule() != Rule::member_initializer {
                            continue;
                        }
                        let mut mi_inner = mi.into_inner();
                        let prop_name = mi_inner.next().unwrap().as_str().to_string();
                        let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                        members.push((prop_name, prop_expr));
                    }
                    if let Some(class_name) = &type_hint {
                        init = Some(emit_vb_object_init_iife(Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident(class_name)),
                            args: ctor_args.clone(),
                        }), members));
                    }
                }
                _ => {}
            }
        }

        if array_rank_count > 0 {
            if let Some(type_hint_value) = type_hint.as_mut() {
                for _ in 0..array_rank_count {
                    type_hint_value.push_str("()");
                }
            }
        }

        if is_new && init.is_none() {
            if let Some(class_name) = &type_hint {
                init = Some(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(class_name)),
                    args: ctor_args,
                }));
            }
        }

        decls.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds,
            with_events: false,
        });
    }
    Ok(decls)
}

fn parse_redim_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut preserve = false;
    let mut array = String::new();
    let mut bounds = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::preserve_keyword => preserve = true,
            Rule::identifier => array = p.as_str().to_string(),
            Rule::array_bounds => {
                bounds = parse_array_bounds_pair(p)?;
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::ReDim {
        preserve,
        array,
        bounds,
    }, span))
}

fn parse_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::dim_statement => {
            let decls = parse_dim_statement(pair)?;
            StmtKind::VarDecl {
                declarations: decls,
                kind: VarDeclKind::Dim,
            }
        }
        Rule::const_statement => {
            let (_vis, decl) = parse_const_statement(pair)?;
            StmtKind::VarDecl {
                declarations: vec![decl],
                kind: VarDeclKind::Const,
            }
        }
        Rule::redim_statement => {
            return parse_redim_statement(pair);
        }
        Rule::erase_statement => {
            let array = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string())
                .ok_or_else(|| "Erase statement missing array name".to_string())?;
            StmtKind::Erase { array }
        }
        Rule::select_statement => {
            return parse_select_statement(pair);
        }
        Rule::dot_assign_statement => {
            // .prop1.prop2 = value (inside With block)
            let inner = pair.into_inner();
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(p.as_str().to_string()),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| "dot_assign missing value".to_string())?;
            if members.is_empty() {
                return Err("dot_assign needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            // Build target: WithTarget.member1.member2...lastMember
            let mut obj = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::me_assign_statement => {
            // Me.prop1.prop2 = value
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(p.as_str().to_string()),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| "me_assign_statement missing value".to_string())?;
            if members.is_empty() {
                return Err("me_assign_statement needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            let mut obj = Expression::new(ExprKind::This);
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::mybase_assign_statement => {
            // MyBase.prop = value
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(p.as_str().to_string()),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| "mybase_assign_statement missing value".to_string())?;
            if members.is_empty() {
                return Err("mybase_assign_statement needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            let mut obj = Expression::new(ExprKind::Super);
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::assign_statement => {
            let mut inner = pair.into_inner();
            // First child is l_value_expression
            let lhs_pair = inner.next().unwrap();
            let lhs_expr = parse_l_value_expression(lhs_pair)?;
            let value_expr = parse_expression(inner.next().unwrap())?;

            StmtKind::Assign {
                targets: vec![lhs_expr],
                value: value_expr,
            }
        }
        Rule::set_statement => {
            let mut inner = pair.into_inner();
            let target_name = inner.next().unwrap().as_str().to_string();
            let value = parse_expression(inner.next().unwrap())?;

            StmtKind::Assign {
                targets: vec![Expression::ident(&target_name)],
                value,
            }
        }
        Rule::compound_assign_statement => {
            let mut inner = pair.into_inner();
            let lhs_pair = inner.next().unwrap();
            let lhs_expr = parse_l_value_expression(lhs_pair)?;

            let op_pair = inner.next().unwrap();
            let op = match op_pair.as_str() {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "\\=" => CompoundOp::IDiv,
                "&=" => CompoundOp::Concat,
                "^=" => CompoundOp::Pow,
                "<<=" => CompoundOp::Shl,
                ">>=" => CompoundOp::Shr,
                _ => return Err(format!("Unknown compound assignment operator: {}", op_pair.as_str())),
            };

            let value = parse_expression(inner.next().unwrap())?;

            StmtKind::CompoundAssign {
                target: lhs_expr,
                op,
                value,
            }
        }
        Rule::raiseevent_statement => {
            let mut inner = pair.into_inner();
            let event_name = inner.next().unwrap().as_str().to_string();
            let mut args = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for arg in p.into_inner() {
                        args.push(parse_expression(arg)?);
                    }
                }
            }
            StmtKind::RaiseEvent { event_name, args }
        }
        Rule::if_statement => return parse_if_statement(pair),
        Rule::single_line_if_statement => return parse_single_line_if(pair),
        Rule::for_each_statement => return parse_for_each_statement(pair),
        Rule::for_statement => return parse_for_statement(pair),
        Rule::while_statement => return parse_while_statement(pair),
        Rule::do_loop_statement => return parse_do_loop_statement(pair),
        Rule::with_statement => return parse_with_statement(pair),
        Rule::using_statement => return parse_using_statement(pair),
        Rule::exit_statement => {
            let mut inner = pair.into_inner();
            let exit_type = inner.next()
                .ok_or_else(|| "Exit statement missing type".to_string())?
                .as_str()
                .to_lowercase();

            match exit_type.as_str() {
                "sub" => StmtKind::Break(BreakTarget::Kind(ExitKind::Sub)),
                "function" => StmtKind::Break(BreakTarget::Kind(ExitKind::Function)),
                "for" => StmtKind::Break(BreakTarget::Kind(ExitKind::For)),
                "do" => StmtKind::Break(BreakTarget::Kind(ExitKind::Do)),
                "while" => StmtKind::Break(BreakTarget::Kind(ExitKind::While)),
                "select" => StmtKind::Break(BreakTarget::Kind(ExitKind::Select)),
                "try" => StmtKind::Break(BreakTarget::Kind(ExitKind::Try)),
                "property" => StmtKind::Break(BreakTarget::Kind(ExitKind::Property)),
                _ => return Err(format!("Unknown exit type: {}", exit_type)),
            }
        }
        Rule::try_statement => return parse_try_statement(pair),
        Rule::throw_statement => {
            let mut inner = pair.into_inner();
            let expr = inner.next().map(parse_expression).transpose()?;
            StmtKind::Throw { expr, cause: None }
        }
        Rule::continue_statement => return parse_continue_statement(pair),
        Rule::open_statement => return parse_open_statement(pair),
        Rule::close_statement => return parse_close_statement(pair),
        Rule::print_file_statement => return parse_print_file_statement(pair),
        Rule::write_file_statement => return parse_write_file_statement(pair),
        Rule::input_file_statement => return parse_input_file_statement(pair),
        Rule::line_input_statement => return parse_line_input_statement(pair),
        Rule::return_statement => {
            let mut inner = pair.into_inner();
            let value = inner.next().map(parse_expression).transpose()?;
            StmtKind::Return(value)
        }
        Rule::call_statement => {
            let mut inner = pair.into_inner();
            let mut first = inner.next().unwrap();

            // Skip optional Call keyword
            if first.as_rule() == Rule::call_keyword {
                first = inner.next().unwrap();
            }

            // Check if it's a member_call, member_access, call_expression, me_member_call, cast_member_call, or simple identifier
            match first.as_rule() {
                Rule::postfix | Rule::cast_member_call | Rule::me_member_call | Rule::mybase_member_call | Rule::member_call | Rule::member_access | Rule::call_expression => {
                    // Parse as expression and convert to statement
                    let expr = parse_expression(first)?;
                    StmtKind::Expr(expr)
                }
                Rule::identifier => {
                    // Could be: identifier, identifier(args), or identifier args
                    let name = first.as_str().to_string();
                    let arguments = inner.next()
                        .map(|p| {
                            if p.as_rule() == Rule::argument_list {
                                parse_argument_list(p)
                            } else {
                                // Single expression argument
                                parse_expression(p).map(|e| vec![Argument::positional(e)])
                            }
                        })
                        .transpose()?
                        .unwrap_or_default();

                    StmtKind::Expr(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(&name)),
                        args: arguments,
                        optional: false,
                    }))
                }
                _ => {
                    let name = first.as_str().to_string();
                    StmtKind::Expr(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(&name)),
                        args: vec![],
                        optional: false,
                    }))
                }
            }
        }
        Rule::expression_statement => {
            let expr = parse_expression(pair.into_inner().next().unwrap())?;
            StmtKind::Expr(expr)
        }
        Rule::addhandler_statement => {
            let mut inner = pair.into_inner();
            let event_target_str = inner.next().unwrap().as_str().to_string();
            let handler = parse_expression(inner.next().unwrap())?;
            let (control, event) = split_event_target(&event_target_str);
            crate::common::events::add_handler_stmt(
                control,
                event,
                handler,
            )
        }
        Rule::removehandler_statement => {
            let mut inner = pair.into_inner();
            let event_target_str = inner.next().unwrap().as_str().to_string();
            let handler = parse_expression(inner.next().unwrap())?;
            let (control, event) = split_event_target(&event_target_str);
            crate::common::events::remove_handler_stmt(
                control,
                event,
                handler,
            )
        }
        Rule::static_statement => {
            let mut name = String::new();
            let mut var_type: Option<String> = None;
            let mut initializer = None;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::type_name => var_type = Some(p.as_str().to_string()),
                    Rule::expression => initializer = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: var_type,
                    init: initializer,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Static,
            }
        }
        Rule::goto_statement => {
            let label = pair.into_inner().next().unwrap().as_str().to_string();
            StmtKind::GoTo(label)
        }
        Rule::label_statement => {
            let label = pair.into_inner().next().unwrap().as_str().to_string();
            StmtKind::Label(label)
        }
        Rule::on_error_statement => {
            let text = pair.as_str().to_lowercase();
            if text.contains("resume") && text.contains("next") {
                StmtKind::OnErrorResumeNext
            } else {
                // On Error GoTo <label> or On Error GoTo 0
                let inner = pair.into_inner();
                let target = inner.last().map(|p| p.as_str().to_string()).unwrap_or_else(|| "0".to_string());
                StmtKind::OnErrorGoTo(target)
            }
        }
        Rule::resume_statement => {
            // Resume → Empty (simplified in common AST)
            StmtKind::Empty
        }
        // New declarations — parse gracefully as no-op statements for now
        Rule::interface_decl | Rule::structure_decl |
        Rule::event_decl | Rule::delegate_sub_decl | Rule::delegate_function_decl => {
            StmtKind::Empty
        }
        Rule::namespace_decl => {
            StmtKind::Empty
        }
        Rule::synclock_statement => return parse_synclock_statement(pair),
        _ => return Err(format!("Unexpected rule: {:?}", pair.as_rule())),
    };
    Ok(Statement::with_span(kind, span))
}

fn parse_if_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;
    let mut then_body = Vec::new();
    let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::if_body => {
                if then_body.is_empty() {
                    then_body = parse_block(p)?;
                }
            }
            Rule::elseif_block => {
                let mut elseif_condition = None;
                let mut elseif_body = Vec::new();
                for p_inner in p.into_inner() {
                    match p_inner.as_rule() {
                        Rule::expression => elseif_condition = Some(parse_expression(p_inner)?),
                        Rule::if_body => { elseif_body = parse_block(p_inner)?; break; }
                        _ => {}
                    }
                }
                if let Some(cond) = elseif_condition {
                    elifs.push((cond, elseif_body));
                }
            }
            Rule::else_block => {
                let mut body = Vec::new();
                for p_inner in p.into_inner() {
                    if p_inner.as_rule() == Rule::if_body {
                        body = parse_block(p_inner)?;
                        break;
                    }
                }
                else_body = Some(body);
            }
            Rule::NEWLINE | Rule::if_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    }, span))
}

fn parse_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut statements = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    statements.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::statement => {
                statements.push(parse_statement(p)?);
            }
            Rule::NEWLINE | Rule::EOI => {}
            _ => {}
        }
    }
    Ok(statements)
}

fn parse_for_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let variable = inner.next().unwrap().as_str().to_string();

    // Skip optional 'As type_name'
    let mut variable_type = None;
    let mut next = inner.next().unwrap();
    if next.as_rule() == Rule::type_name {
        variable_type = Some(next.as_str().to_string());
        next = inner.next().unwrap();
    }
    let start = parse_expression(next)?;
    let end = parse_expression(inner.next().unwrap())?;

    let mut step = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::expression => step = Some(parse_expression(p)?),
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::NEWLINE | Rule::for_end => {}
            _ => body.push(parse_statement(p)?),
        }
    }

    // Convert VB For to C-style For:
    // init: Dim variable = start
    // cond: variable <= end (or >= end if step is negative)
    // update: variable = variable + step
    let init = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(variable.clone()),
            type_hint: variable_type,
            init: Some(start),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Dim,
    });

    let step_val = step.unwrap_or_else(|| Expression::int(1));
    // VB ranges count up by default but `Step -N` reverses the
    // direction; the loop condition flips accordingly. Detect a
    // negative literal step and emit `i >= end` instead of
    // `i <= end`. Non-literal step expressions (rare) keep the
    // up-counting semantics — runtime evaluation can't pick a
    // direction at compile time without a helper.
    let step_is_negative = match &step_val.kind {
        ExprKind::Lit(Literal::Int(n)) => *n < 0,
        ExprKind::Lit(Literal::Float(f)) => *f < 0.0,
        ExprKind::Unary { op: UnaryOp::Neg, .. } => true,
        _ => false,
    };
    let cond_op = if step_is_negative { BinOp::GtEq } else { BinOp::LtEq };
    let cond = Expression::new(ExprKind::Binary {
        op: cond_op,
        left: Box::new(Expression::ident(&variable)),
        right: Box::new(end),
    });
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&variable)),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&variable)),
            right: Box::new(step_val),
        })),
    });

    Ok(Statement::with_span(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update),
        body,
    }, span))
}

fn parse_while_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::NEWLINE | Rule::while_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::While {
        cond,
        body,
        else_body: None,
    }, span))
}

fn parse_do_loop_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut pre_condition: Option<(bool, Expression)> = None; // (is_until, expr)
    let mut post_condition: Option<(bool, Expression)> = None;
    let mut body = Vec::new();
    let mut current_is_until = false;

    for p in inner {
        match p.as_rule() {
            Rule::do_while_kw => current_is_until = false,
            Rule::do_until_kw => current_is_until = true,
            Rule::expression => {
                // Determine if it's pre or post condition based on position
                if body.is_empty() {
                    pre_condition = Some((current_is_until, parse_expression(p)?));
                } else {
                    post_condition = Some((current_is_until, parse_expression(p)?));
                }
            }
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::do_end => {
                // Parse post-condition from do_end children (Loop While/Until)
                for dp in p.into_inner() {
                    match dp.as_rule() {
                        Rule::do_while_kw => current_is_until = false,
                        Rule::do_until_kw => current_is_until = true,
                        Rule::expression => {
                            post_condition = Some((current_is_until, parse_expression(dp)?));
                        }
                        _ => {}
                    }
                }
            }
            Rule::NEWLINE => {}
            _ => {}
        }
    }

    // Map to common AST:
    // If there's a pre_condition: it's a While loop (with condition potentially inverted for Until)
    // If there's a post_condition: it's a DoWhile loop
    // If neither: infinite loop (DoWhile with true condition)
    if let Some((is_until, cond)) = pre_condition {
        // Do While/Until <cond> ... Loop → While loop
        let effective_cond = if is_until {
            Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(cond) })
        } else {
            cond
        };
        Ok(Statement::with_span(StmtKind::While {
            cond: effective_cond,
            body,
            else_body: None,
        }, span))
    } else if let Some((is_until, cond)) = post_condition {
        // Do ... Loop While/Until <cond> → DoWhile
        Ok(Statement::with_span(StmtKind::DoWhile {
            body,
            cond,
            until: is_until,
        }, span))
    } else {
        // Do ... Loop (infinite)
        Ok(Statement::with_span(StmtKind::DoWhile {
            body,
            cond: Expression::bool(true),
            until: false,
        }, span))
    }
}

fn parse_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut pair = pair;
    loop {
        let span = to_span(&pair);
        let kind = match pair.as_rule() {
            Rule::argument => {
                let inner = pair.into_inner().next().unwrap();
                pair = inner;
                continue;
            }
            Rule::named_argument => {
                let mut inner = pair.into_inner();
                let _name = inner.next();
                pair = inner.next().ok_or_else(|| "Named argument missing value".to_string())?;
                continue;
            }
            Rule::expression | Rule::logical_xor | Rule::logical_or | Rule::logical_and |
            Rule::equality | Rule::comparison | Rule::bit_shift | Rule::additive |
            Rule::multiplicative | Rule::exponent => {
                let mut probe = pair.clone().into_inner();
                let first = probe.next().unwrap();
                if probe.next().is_none() {
                    pair = first;
                    continue;
                }
                return parse_binary_expression(pair);
            }
            Rule::not_condition => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                if first.as_rule() == Rule::not_op {
                    let operand = parse_expression(inner.next().unwrap())?;
                    ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(operand),
                    }
                } else {
                    pair = first;
                    continue;
                }
            }
            Rule::lambda_expression => return parse_lambda_expression(pair),
            Rule::nameof_expression => {
                let name = pair.into_inner()
                    .find(|p| p.as_rule() == Rule::dotted_identifier)
                    .map(|p| {
                        let text = p.as_str();
                        text.rsplit('.').next().unwrap_or(text).to_string()
                    })
                    .unwrap_or_default();
                ExprKind::Lit(Literal::Str(name))
            }
            Rule::gettype_expression => {
                let type_name = pair.into_inner()
                    .find(|p| p.as_rule() == Rule::type_name)
                    .map(|p| p.as_str().trim().to_string())
                    .unwrap_or_default();
                ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(type_name.rsplit('.').next().unwrap_or(&type_name)),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("FullName"),
                        value: Expression::string(&type_name),
                    },
                ])
            }
            Rule::typeof_expression => {
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().trim().to_string();
                ExprKind::IsType {
                    expr: Box::new(expr),
                    type_name,
                }
            }
            Rule::unary => return parse_unary_expression(pair),
            Rule::postfix => return parse_postfix_expression(pair),
            Rule::call_expression => {
                let mut inner = pair.into_inner();
                let name = inner.next().unwrap().as_str().to_string();
                let arguments = inner.next()
                    .map(parse_argument_list)
                    .transpose()?
                    .unwrap_or_default();

                if let Some(rewritten) = canonicalize_call(&name, &arguments) {
                    return Ok(rewritten);
                }

                ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::member_call => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }

                return Ok(expr);
            }
            Rule::member_access => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for p in inner {
                    expr = canonicalize_member_access(expr, &normalize_vb_identifier(p.as_str()));
                }

                return Ok(expr);
            }
            Rule::query_expression => return parse_query_expression(pair),
            Rule::xml_literal => return parse_xml_literal(pair),
            Rule::identifier => ExprKind::Ident(normalize_vb_identifier(pair.as_str())),
            Rule::cast_expression => {
                let text = pair.as_str();
                let cast_kind = if text.len() >= 10 && text[..10].eq_ignore_ascii_case("DirectCast") {
                    "DirectCast"
                } else if text.len() >= 7 && text[..7].eq_ignore_ascii_case("TryCast") {
                    "TryCast"
                } else {
                    "CType"
                };
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().to_string();
                let full_type = if cast_kind != "CType" {
                    format!("{}:{}", cast_kind, type_name)
                } else {
                    type_name
                };
                ExprKind::Cast {
                    expr: Box::new(expr),
                    type_name: full_type,
                }
            }
            Rule::cast_member_call => {
                let mut inner = pair.into_inner();
                let cast_pair = inner.next().unwrap();
                let mut expr = parse_expression(cast_pair)?;
                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }
                return Ok(expr);
            }
            Rule::interpolated_string => {
                let s = pair.as_str();
                let inner_str = s[2..s.len() - 1].replace("\"\"", "\"");
                let mut parts = Vec::new();
                let mut current_text = String::new();
                let mut chars = inner_str.chars().peekable();

                while let Some(ch) = chars.next() {
                    if ch == '{' {
                        if chars.peek() == Some(&'{') {
                            chars.next();
                            current_text.push('{');
                            continue;
                        }
                        if !current_text.is_empty() {
                            parts.push(InterpolPart::Text(current_text.clone()));
                            current_text.clear();
                        }
                        let mut expr_text = String::new();
                        let mut depth = 1;
                        while let Some(c) = chars.next() {
                            if c == '{' {
                                depth += 1;
                            }
                            if c == '}' {
                                depth -= 1;
                                if depth == 0 {
                                    break;
                                }
                            }
                            expr_text.push(c);
                        }
                        match parse_expression_str(&expr_text) {
                            Ok(expr) => parts.push(InterpolPart::Expr(expr)),
                            Err(_) => parts.push(InterpolPart::Expr(Expression::ident(expr_text.trim()))),
                        }
                    } else if ch == '}' {
                        if chars.peek() == Some(&'}') {
                            chars.next();
                            current_text.push('}');
                        }
                    } else {
                        current_text.push(ch);
                    }
                }

                if !current_text.is_empty() {
                    parts.push(InterpolPart::Text(current_text));
                }

                if parts.is_empty() {
                    ExprKind::Lit(Literal::Str(String::new()))
                } else if parts.len() == 1 {
                    match parts.into_iter().next().unwrap() {
                        InterpolPart::Text(s) => ExprKind::Lit(Literal::Str(s)),
                        InterpolPart::Expr(expr) => return Ok(expr),
                        InterpolPart::Formatted(expr, _) => return Ok(expr),
                    }
                } else {
                    ExprKind::Interpolation(parts)
                }
            }
            Rule::string_literal => {
                let s = pair.as_str();
                ExprKind::Lit(Literal::Str(s[1..s.len() - 1].replace("\"\"", "\"")))
            }
            Rule::numeric_literal => {
                let s = pair.as_str();
                let s = s.trim_end_matches(|c: char| {
                    c.is_ascii_alphabetic() || c == '!' || c == '#' || c == '@' || c == '%'
                });
                if s.contains('.') {
                    ExprKind::Lit(Literal::Float(s.parse().unwrap_or(0.0)))
                } else {
                    ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0)))
                }
            }
            Rule::boolean_literal => ExprKind::Lit(Literal::Bool(pair.as_str().eq_ignore_ascii_case("true"))),
            Rule::array_literal => {
                let elements = pair.into_inner()
                    .map(|p| parse_expression(p).map(|value| ArrayElement {
                        key: None,
                        value,
                        spread: false,
                        by_ref: false,
                    }))
                    .collect::<Result<Vec<_>, _>>()?;
                ExprKind::Array(elements)
            }
            Rule::date_literal => {
                let s = pair.as_str();
                ExprKind::Lit(Literal::Str(s[1..s.len() - 1].trim().to_string()))
            }
            Rule::nothing_literal => ExprKind::Lit(Literal::Null),
            Rule::new_expression => {
                let mut inner = pair.into_inner();
                let id_pair = inner.next().unwrap();
                let mut class_name = id_pair.as_str().to_string();
                let mut args = Vec::new();
                let mut array_init: Option<Vec<Expression>> = None;
                for p in inner {
                    match p.as_rule() {
                        Rule::generic_suffix => class_name.push_str(p.as_str()),
                        Rule::argument_list => args = parse_argument_list(p)?,
                        Rule::array_literal => {
                            let elements = p.into_inner()
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            array_init = Some(elements);
                        }
                        Rule::from_initializer => {
                            let elements = p.into_inner()
                                .filter(|e| e.as_rule() == Rule::expression)
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            let mut all_args = args;
                            for elem in elements {
                                all_args.push(Argument::positional(elem));
                            }
                            return Ok(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args: all_args,
                            }, span));
                        }
                        Rule::with_initializer => {
                            let mut members = Vec::new();
                            for mi in p.into_inner() {
                                if mi.as_rule() != Rule::member_initializer {
                                    continue;
                                }
                                let mut mi_inner = mi.into_inner();
                                let prop_name = mi_inner.next().unwrap().as_str().to_string();
                                let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                                members.push((prop_name, prop_expr));
                            }
                            return Ok(emit_vb_object_init_iife(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args,
                            }, span), members));
                        }
                        _ => {}
                    }
                }
                if let Some(elements) = array_init {
                    ExprKind::Array(elements.into_iter().map(|value| ArrayElement {
                        key: None,
                        value,
                        spread: false,
                        by_ref: false,
                    }).collect())
                } else {
                    ExprKind::New {
                        class: Box::new(Expression::ident(&class_name)),
                        args,
                    }
                }
            }
            Rule::if_expression => {
                let mut inner = pair.into_inner();
                let first = parse_expression(inner.next().unwrap())?;
                let second = parse_expression(inner.next().unwrap())?;
                let third = inner.next().map(parse_expression).transpose()?;
                match third {
                    Some(else_expr) => ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    },
                    None => ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    },
                }
            }
            Rule::addressof_expr => {
                let mut name = String::new();
                for p in pair.into_inner() {
                    if p.as_rule() == Rule::dotted_identifier {
                        name = p.as_str().to_string();
                    }
                }
                ExprKind::AddressOf(name)
            }
            Rule::me_keyword => ExprKind::This,
            Rule::dot_call_statement => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("dot_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::dot_member_access => {
                let inner = pair.into_inner();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_access => {
                let mut inner = pair.into_inner();
                let _me = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::This);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::mybase_member_access => {
                let mut inner = pair.into_inner();
                let _mybase = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::Super);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::me_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("me_member_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::This);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::mybase_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::mybase_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("mybase_member_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Super);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                if identifiers.len() == 1 {
                    ExprKind::SuperCall {
                        method: Some(method_name),
                        args: arguments,
                    }
                } else {
                    let callee = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: method_name,
                        null_safe: false,
                    });
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args: arguments,
                        optional: false,
                    }
                }
            }
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };
        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "and" | "andalso" => BinOp::And,
                    "or" | "orelse" => BinOp::Or,
                    "xor" => BinOp::Xor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left),
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}

/*
fn parse_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut pair = pair;
    loop {
        let span = to_span(&pair);
        let kind = match pair.as_rule() {
            Rule::expression | Rule::logical_xor | Rule::logical_or | Rule::logical_and |
            Rule::equality | Rule::comparison | Rule::bit_shift | Rule::additive |
            Rule::multiplicative | Rule::exponent => {
                let mut probe = pair.clone().into_inner();
                let first = probe.next().unwrap();
                if probe.next().is_none() {
                    pair = first;
                    continue;
                }
                return parse_binary_expression(pair);
            }
            Rule::not_condition => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                if first.as_rule() == Rule::not_op {
                    let operand = parse_expression(inner.next().unwrap())?;
                    ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(operand) }
                } else {
                    pair = first;
                    continue;
                }
            }
            Rule::lambda_expression => {
                return parse_lambda_expression(pair);
            }
            Rule::typeof_expression => {
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().trim().to_string();
                ExprKind::IsType {
                    expr: Box::new(expr),
                    type_name,
                }
            }
            Rule::unary => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                match first.as_rule() {
                    Rule::neg_op => {
                        let operand = parse_expression(inner.next().unwrap())?;
                        ExprKind::Unary {
                            op: UnaryOp::Neg,
                            expr: Box::new(operand),
                        }
                    }
                    Rule::await_op => {
                        let operand = parse_expression(inner.next().unwrap())?;
                        ExprKind::Await(Box::new(operand))
                    }
                    _ => {
                        pair = first;
                        continue;
                    }
                }
            }
            Rule::postfix => {
                let mut inner = pair.into_inner();
                let primary = inner.next().unwrap();
                let Some(first_chain) = inner.next() else {
                    pair = primary;
                    continue;
                };

                let mut expr = parse_expression(primary)?;
                expr = parse_member_chain_node(first_chain, expr)?;
                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }
                return Ok(expr);
            }
            Rule::call_expression => {
                let mut inner = pair.into_inner();
                let name = inner.next().unwrap().as_str().to_string();
                let arguments = inner.next()
                    .map(parse_argument_list)
                    .transpose()?
                    .unwrap_or_default();

                if let Some(rewritten) = canonicalize_call(&name, &arguments) {
                    return Ok(rewritten);
                }

                ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::member_call => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }

                return Ok(expr);
            }
            Rule::member_access => {

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let _span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "and" | "andalso" => BinOp::And,
                    "or" | "orelse" => BinOp::Or,
                    "xor" => BinOp::Xor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left),
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}
                    }
                } else if ch == '}' {
                    // Check for }} escape (literal brace)
                    if chars.peek() == Some(&'}') {
                        chars.next();
                        current_text.push('}');
                    }
                } else {
                    current_text.push(ch);
                }
            }
            // Flush remaining text
            if !current_text.is_empty() {
                parts.push(InterpolPart::Text(current_text));
            }

            if parts.is_empty() {
                ExprKind::Lit(Literal::Str(String::new()))
            } else if parts.len() == 1 {
                match parts.into_iter().next().unwrap() {
                    InterpolPart::Text(s) => ExprKind::Lit(Literal::Str(s)),
                    InterpolPart::Expr(e) => return Ok(e),
                    InterpolPart::Formatted(e, _) => return Ok(e),
                }
            } else {
                ExprKind::Interpolation(parts)
            }
        }
        Rule::string_literal => {
            let s = pair.as_str();
            // Strip outer quotes, then unescape VB-style doubled quotes ("" -> ")
            let inner = s[1..s.len()-1].replace("\"\"", "\"");
            ExprKind::Lit(Literal::Str(inner))
        }
        Rule::numeric_literal => {
            let s = pair.as_str();
            // Strip type suffixes: F, D, L, R, S, US, UI, UL, !, #, @, %
            let s = s.trim_end_matches(|c: char| c.is_ascii_alphabetic() || c == '!' || c == '#' || c == '@' || c == '%');
            if s.contains('.') {
                ExprKind::Lit(Literal::Float(s.parse().unwrap_or(0.0)))
            } else {
                ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0)))
            }
        }
        Rule::boolean_literal => {
            ExprKind::Lit(Literal::Bool(pair.as_str().to_lowercase() == "true"))
        }
        Rule::array_literal => {
            return parse_array_literal(pair);
        }
        Rule::date_literal => {
            let s = pair.as_str();
            // Strip the surrounding # delimiters
            let inner = s[1..s.len()-1].trim().to_string();
            ExprKind::Lit(Literal::Str(inner))
        }
        Rule::nothing_literal => ExprKind::Lit(Literal::Null),
        Rule::new_expression => {
            let mut inner = pair.into_inner();
            let id_pair = inner.next().unwrap();
            let mut class_name = id_pair.as_str().to_string();
            // Check for generic_suffix: List(Of String) -> "List(Of String)"
            let mut args: Vec<Argument> = Vec::new();
            let mut array_init: Option<Vec<Expression>> = None;
            for p in inner {
                match p.as_rule() {
                    Rule::generic_suffix => {
                        class_name.push_str(p.as_str());
                    }
                    Rule::argument_list => {
                        args = parse_argument_list(p)?;
                    }
                    Rule::array_literal => {
                        // New Type() {elem1, elem2, ...} → array initializer
                        let elements: Vec<Expression> = p.into_inner()
                            .map(|e| parse_expression(e))
                            .collect::<Result<Vec<_>, _>>()?;
                        array_init = Some(elements);
                    }
                    Rule::from_initializer => {
                        // New List(Of T) From { expr, expr, ... }
                        let elements: Vec<Expression> = p.into_inner()
                            .filter(|e| e.as_rule() == Rule::expression)
                            .map(|e| parse_expression(e))
                            .collect::<Result<Vec<_>, _>>()?;
                        let mut all_args = args;
                        for elem in elements {
                            all_args.push(Argument::positional(elem));
                        }
                        return Ok(Expression::with_span(ExprKind::New {
                            class: Box::new(Expression::ident(&class_name)),
                            args: all_args,
                        }, span));
                    }
                    Rule::with_initializer => {
                        // New Type() With { .Prop = expr, ... }
                        let mut members = Vec::new();
                        for mi in p.into_inner() {
                            if mi.as_rule() != Rule::member_initializer { continue; }
                            let mut mi_inner = mi.into_inner();
                            let prop_name = mi_inner.next().unwrap().as_str().to_string();
                            let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                            members.push((prop_name, prop_expr));
                        }
                        return Ok(emit_vb_object_init_iife(Expression::with_span(ExprKind::New {
                            class: Box::new(Expression::ident(&class_name)),
                            args,
                        }, span), members));
                    }
                    _ => {}
                }
            }
            // If there's an array initializer, return an Array instead of New
            if let Some(elements) = array_init {
                ExprKind::Array(
                    elements.into_iter().map(|e| ArrayElement {
                        key: None,
                        value: e,
                        spread: false,
                        by_ref: false,
                    }).collect()
                )
            } else {
                ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args,
                }
            }
        }
        Rule::if_expression => {
            let mut inner = pair.into_inner();
            let first = parse_expression(inner.next().unwrap())?;
            let second = parse_expression(inner.next().unwrap())?;
            let third = inner.next().map(|p| parse_expression(p)).transpose()?;
            match third {
                Some(else_expr) => {
                    ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    }
                }
                None => {
                    // If(a, b) with no else → null coalesce
                    ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    }
                }
            }
        }
        Rule::addressof_expr => {
            let inner = pair.into_inner();
            let mut name = String::new();
            for p in inner {
                if p.as_rule() == Rule::dotted_identifier {
                    name = p.as_str().to_string();
                }
            }
            ExprKind::AddressOf(name)
        }
        Rule::me_keyword => {
            ExprKind::This
        }
        Rule::dot_call_statement => {
            // .Method(args) or .obj.Method(args) inside With block
            let inner = pair.into_inner();
            let mut identifiers = Vec::new();
            let mut arguments: Vec<Argument> = Vec::new();
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }
            if identifiers.is_empty() {
                return Err("dot_call needs at least one identifier".to_string());
            }
            let method_name = identifiers.last().unwrap().clone();
            let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }
            // Build callee as member access, then call it
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe: false,
            });
            ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }
        }
        Rule::dot_member_access => {
            // .prop or .obj.prop inside With block
            let inner = pair.into_inner();
            let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::me_member_access => {
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut expr = Expression::new(ExprKind::This);
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::mybase_member_access => {
            // MyBase.Property
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut expr = Expression::new(ExprKind::Super);
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::me_member_call => {
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments: Vec<Argument> = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::me_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err("me_member_call needs at least one identifier".to_string());
            }

            // Last identifier is the method name
            let method_name = identifiers.last().unwrap().clone();

            // Build object expression: Me.a.b... (all except last)
            let mut expr = Expression::new(ExprKind::This);
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }

            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe: false,
            });
            ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }
        }
        Rule::mybase_member_call => {
            // MyBase.Method()
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments: Vec<Argument> = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::mybase_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err("mybase_member_call needs at least one identifier".to_string());
            }

            let method_name = identifiers.last().unwrap().clone();
            let mut expr = Expression::new(ExprKind::Super);
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }

            // MyBase.Method(args) → SuperCall
            if identifiers.len() == 1 {
                ExprKind::SuperCall {
                    method: Some(method_name),
                    args: arguments,
                }
            } else {
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
        }
        _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
    };
    Ok(Expression::with_span(kind, span))
                ExprKind::Lit(Literal::Str(format!("{}{}", fmt, arg_expr)))
            }
            Rule::date_literal => {
                let s = pair.as_str();
                let inner = s.trim_matches('#').to_string();
                ExprKind::Lit(Literal::Str(inner))
            }
            Rule::string_literal => {
                let raw = pair.as_str();
                let inner = &raw[1..raw.len()-1];
                let s = inner.replace("\"\"", "\"");
                ExprKind::Lit(Literal::Str(s))
            }
            Rule::numeric_literal => {
                let text = pair.as_str();
                let trimmed = text.trim_end_matches(|c: char| {
                    c == 'F' || c == 'f' || c == 'D' || c == 'd' || c == 'L' || c == 'l' ||
                    c == 'R' || c == 'r' || c == 'S' || c == 's' || c == '!' || c == '#' ||
                    c == '@' || c == '%' || c == 'U' || c == 'u' || c == 'I' || c == 'i'
                });
                if trimmed.contains('.') {
                    ExprKind::Lit(Literal::Float(trimmed.parse().unwrap_or(0.0)))
                } else {
                    ExprKind::Lit(Literal::Int(trimmed.parse().unwrap_or(0)))
                }
            }
            Rule::boolean_literal => {
                ExprKind::Lit(Literal::Bool(pair.as_str().eq_ignore_ascii_case("true")))
            }
            Rule::nothing_literal => ExprKind::Lit(Literal::Null),
            Rule::array_literal => {
                let elements = pair.into_inner()
                    .map(|p| parse_expression(p).map(ArrayElement::value))
                    .collect::<Result<Vec<_>, _>>()?;
                ExprKind::Array(elements)
            }
            Rule::new_expression => return parse_new_expression(pair),
            Rule::if_expression => return parse_if_expression(pair),
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };
        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let _span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "and" | "andalso" => BinOp::And,
                    "or" | "orelse" => BinOp::Or,
                    "xor" => BinOp::Xor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left), // Should not happen with current grammar
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}

*/
fn parse_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();

    match first.as_rule() {
        Rule::not_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            }, span))
        }
        Rule::neg_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }, span))
        }
        Rule::await_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(ExprKind::Await(Box::new(operand)), span))
        }
        Rule::postfix => {
            parse_postfix_expression(first)
        }
        _ => {
            // Fallback: treat as primary
            parse_expression(first)
        }
    }
}

fn parse_postfix_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let primary = inner.next().unwrap();
    let mut expr = parse_expression(primary)?;

    // Apply member_chain postfix operations
    for chain in inner {
        expr = parse_member_chain_node(chain, expr)?;
    }

    Ok(expr)
}

fn parse_member_chain_node(chain: Pair<Rule>, expr: Expression) -> Result<Expression, String> {
    match chain.as_rule() {
        Rule::member_chain_invoke => {
            let arguments = chain
                .into_inner()
                .next()
                .map(parse_argument_list)
                .transpose()?
                .unwrap_or_default();
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args: arguments,
                optional: false,
            }))
        }
        Rule::member_chain_call => {
            let mut chain_inner = chain.into_inner();
            let name = chain_inner.next().unwrap().as_str().to_string();
            let arguments = if let Some(arg_list) = chain_inner.next() {
                parse_argument_list(arg_list)?
            } else {
                vec![]
            };
            if name.eq_ignore_ascii_case("Item") && !arguments.is_empty() {
                let mut indexed = expr;
                for arg in arguments {
                    indexed = Expression::new(ExprKind::Index {
                        object: Box::new(indexed),
                        index: Box::new(arg.value),
                        null_safe: false,
                    });
                }
                return Ok(indexed);
            }
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: name,
                null_safe: false,
            });
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }))
        }
        Rule::member_chain_access => {
            let name = chain.into_inner().next().unwrap().as_str().to_string();
            Ok(Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: name,
                null_safe: false,
            }))
        }
        Rule::member_chain => {
            let inner_chain = chain.into_inner().next().unwrap();
            parse_member_chain_node(inner_chain, expr)
        }
        _ => Ok(expr),
    }
}

fn parse_argument_list(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .map(|p| match p.as_rule() {
            Rule::named_argument => {
                let mut inner = p.into_inner();
                let name = normalize_vb_identifier(inner.next().unwrap().as_str());
                let value = parse_expression(inner.next().unwrap())?;
                Ok(Argument {
                    value,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                })
            }
            Rule::argument => {
                let inner = p.into_inner().next().unwrap();
                match inner.as_rule() {
                    Rule::named_argument => {
                        let mut named_inner = inner.into_inner();
                        let name = normalize_vb_identifier(named_inner.next().unwrap().as_str());
                        let value = parse_expression(named_inner.next().unwrap())?;
                        Ok(Argument {
                            value,
                            name: Some(name),
                            by_ref: false,
                            spread: false,
                        })
                    }
                    _ => parse_expression(inner).map(Argument::positional),
                }
            }
            _ => parse_expression(p).map(Argument::positional),
        })
        .collect()
}

fn parse_try_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();

    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for p in inner {
        match p.as_rule() {
            Rule::try_body => {
                body = parse_block_body(p)?;
            }
            Rule::catch_block => catches.push(parse_catch_block(p)?),
            Rule::finally_block => {
                let f_inner = p.into_inner();
                for fp in f_inner {
                    if fp.as_rule() == Rule::try_body {
                        finally = Some(parse_block_body(fp)?);
                    }
                }
            }
            Rule::try_end => {},
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    }, span))
}

fn parse_catch_block(pair: Pair<Rule>) -> Result<CatchClause, String> {
    let inner = pair.into_inner();
    let mut var_name: Option<String> = None;
    let mut catch_type: Option<String> = None;
    let mut when_clause = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                var_name = Some(p.as_str().to_string());
            }
            Rule::type_name => {
                catch_type = Some(p.as_str().to_string());
            }
            Rule::expression => {
                when_clause = Some(parse_expression(p)?);
            }
            Rule::try_body => {
                body = parse_block_body(p)?;
            }
            _ => {}
        }
    }

    Ok(CatchClause {
        types: catch_type.into_iter().collect(),
        var_name,
        stack_var: None,
        body,
        when_clause,
    })
}

fn parse_continue_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let text = pair.as_str().to_lowercase();
    let target = if text.contains("do") {
        ContinueTarget::Kind(ContinueKind::Do)
    } else if text.contains("for") {
        ContinueTarget::Kind(ContinueKind::For)
    } else {
        ContinueTarget::Kind(ContinueKind::While)
    };

    Ok(Statement::with_span(StmtKind::Continue(target), span))
}

fn parse_lambda_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let text = pair.as_str().trim_start();
    let _is_function = text.to_lowercase().starts_with("function");

    let mut inner = pair.into_inner();
    let mut params = Vec::new();

    let mut next_pair = inner.next().ok_or_else(|| "Lambda missing body".to_string())?;

    if next_pair.as_rule() == Rule::param_list {
        params = parse_param_list(next_pair)?;
        next_pair = inner.next().ok_or_else(|| "Lambda missing body".to_string())?;
    }

    let body = match next_pair.as_rule() {
        Rule::expression => LambdaBody::Expr(Box::new(parse_expression(next_pair)?)),
        Rule::NEWLINE => {
            // Multiline block
            let mut body_stmts = Vec::new();
            for item in inner {
                match item.as_rule() {
                    Rule::statement_line => {
                        for stmt_pair in item.into_inner() {
                            if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                                body_stmts.push(parse_statement(stmt_pair)?);
                            }
                        }
                    }
                    _ => {
                        if let Some(decl_stmt) = try_parse_declaration(item.clone())? {
                            body_stmts.push(decl_stmt);
                        }
                    }
                }
            }
            LambdaBody::Block(body_stmts)
        }
        _ => {
            // Any statement variant rule (call_statement, assign_statement, etc.)
            // These appear directly because `statement` is a silent rule in the grammar.
            let stmt = parse_statement(next_pair)?;
            LambdaBody::Block(vec![stmt])
        }
    };

    Ok(Expression::with_span(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: vec![],
    }, span))
}

fn parse_block_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt_pair in pair.into_inner() {
        if stmt_pair.as_rule() == Rule::statement_line {
            for s in stmt_pair.into_inner() {
                if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                    body.push(parse_statement(s)?);
                }
            }
        }
    }
    Ok(body)
}

fn parse_for_each_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let hidden_suffix = pair.as_span().start();
    let mut inner = pair.into_inner();
    let variable = inner.next().unwrap().as_str().to_string();
    let mut variable_type = None;
    let mut collection = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_name => variable_type = Some(p.as_str().to_string()),
            Rule::expression => {
                if collection.is_none() {
                    collection = Some(parse_expression(p)?);
                }
            }
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::for_end => {}
            _ => body.push(parse_statement(p)?),
        }
    }

    let mut loop_var = variable.clone();
    if let Some(type_hint) = variable_type {
        let source_var = format!("__vb_foreach_item_{}", hidden_suffix);
        body.insert(0, Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(variable),
                type_hint: Some(type_hint),
                init: Some(Expression::ident(&source_var)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Dim,
        }));
        loop_var = source_var;
    }

    Ok(Statement::with_span(StmtKind::ForIn {
        var: loop_var,
        key: None,
        iter: collection.ok_or_else(|| "For Each missing collection".to_string())?,
        body,
        of: true, // VB For Each iterates values, like JS for...of
        else_body: None,
        is_async: false,
    }, span))
}

fn parse_with_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let object = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::with_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::With {
        items: vec![WithItem { expr: object, var: None }],
        body,
        is_async: false,
    }, span))
}

fn parse_using_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let variable = inner.next().unwrap().as_str().to_string();

    // Collect remaining pairs
    let remaining: Vec<_> = inner.collect();

    // Find the expression (resource)
    let mut resource_expr = None;
    let mut body_start_idx = 0;

    for (idx, p) in remaining.iter().enumerate() {
        match p.as_rule() {
            Rule::type_name => {}, // Skip type annotation
            Rule::new_expression => {
                resource_expr = Some(parse_expression(p.clone())?);
                body_start_idx = idx + 1;
                break;
            }
            Rule::expression => {
                resource_expr = Some(parse_expression(p.clone())?);
                body_start_idx = idx + 1;
                break;
            }
            _ => {}
        }
    }

    let resource = resource_expr.ok_or_else(|| "Using statement missing resource expression".to_string())?;

    // Parse body statements
    let mut body = Vec::new();
    for p in remaining.iter().skip(body_start_idx) {
        match p.as_rule() {
            Rule::statement_line => {
                for stmt_pair in p.clone().into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::using_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::Using { var: variable, resource, body }, span))
}

fn parse_enum_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut members = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                let text = p.as_str().to_lowercase();
                match text.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Internal,
                    _ => name = p.as_str().to_string(),
                }
            }
            Rule::enum_member | Rule::enum_member_inline => {
                let mut member_inner = p.into_inner();
                let member_name = member_inner.next().unwrap().as_str().to_string();
                let value = member_inner
                    .find(|e| e.as_rule() == Rule::expression)
                    .map(|e| parse_expression(e))
                    .transpose()?;
                members.push(EnumMember { name: member_name, value, constructor_args: Vec::new() });
            }
            Rule::enum_end | Rule::NEWLINE => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::EnumDecl {
        name,
        members,
        visibility,
        is_flags: false,
        backing_type: None,
        interfaces: Vec::new(),
        body_members: Vec::new(),
        decorators: vec![],
    }, span))
}

fn parse_single_line_if(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;

    let then_body = vec![parse_statement(inner.next().unwrap())?];

    let else_body = if let Some(else_body_pair) = inner.next() {
        Some(vec![parse_statement(else_body_pair)?])
    } else {
        None
    };

    Ok(Statement::with_span(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    }, span))
}

fn parse_field_decl(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut field_name = String::new();
    let mut field_type: Option<String> = None;
    let mut field_init = None;
    let mut field_bounds = None;
    let mut is_new = false;
    let mut ctor_args: Vec<Argument> = Vec::new();
    let mut is_with_events = false;

    for fp in pair.into_inner() {
        match fp.as_rule() {
            Rule::withevents_keyword => { is_with_events = true; }
            Rule::visibility_modifier | Rule::sub_modifier_keyword | Rule::partial_keyword => {} // modifiers handled by caller
            Rule::dim_new_keyword => { is_new = true; }
            Rule::identifier => field_name = fp.as_str().to_string(),
            Rule::type_name => field_type = Some(fp.as_str().to_string()),
            Rule::array_rank_spec => {
                field_bounds = Some(parse_array_bounds_pair(fp)?);
            }
            Rule::array_bounds => {
                field_bounds = Some(parse_array_bounds_pair(fp)?);
            }
            Rule::argument_list => {
                for arg_pair in fp.into_inner() {
                    if arg_pair.as_rule() == Rule::expression {
                        ctor_args.push(Argument::positional(parse_expression(arg_pair)?));
                    }
                }
            }
            Rule::expression => field_init = Some(parse_expression(fp)?),
            Rule::array_literal => field_init = Some(parse_array_literal(fp)?),
            _ => {}
        }
    }

    // Handle "As New Type" syntax
    if is_new && field_init.is_none() {
        if let Some(t) = &field_type {
            field_init = Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(t)),
                args: ctor_args,
            }));
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(field_name),
        type_hint: field_type,
        init: field_init,
        array_bounds: field_bounds,
        with_events: is_with_events,
    })
}

fn parse_open_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_path = parse_expression(inner.next().unwrap())?;
    let mode_pair = inner.next().unwrap(); // file_mode
    let mode = match mode_pair.as_str().to_lowercase().as_str() {
        "input" => FileMode::Input,
        "output" => FileMode::Output,
        "append" => FileMode::Append,
        "binary" => FileMode::Binary,
        "random" => FileMode::Random,
        _ => return Err(format!("Unknown file mode: {}", mode_pair.as_str())),
    };
    let file_number = parse_expression(inner.next().unwrap())?;
    Ok(Statement::with_span(StmtKind::OpenFile { path: file_path, mode, file_number }, span))
}

fn parse_close_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    // Close with no arguments closes all files
    let file_number = inner.next().map(|p| parse_expression(p)).transpose()?;
    Ok(Statement::with_span(StmtKind::CloseFile(file_number), span))
}

fn parse_print_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner.next()
        .map(|p| parse_argument_list(p).map(|args| args.into_iter().map(|a| a.value).collect()))
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::with_span(StmtKind::PrintFile { file_number, items }, span))
}

fn parse_write_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner.next()
        .map(|p| parse_argument_list(p).map(|args| args.into_iter().map(|a| a.value).collect()))
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::with_span(StmtKind::WriteFile { file_number, items }, span))
}

fn parse_input_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let mut variables = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::identifier {
            variables.push(p.as_str().to_string());
        }
    }
    Ok(Statement::with_span(StmtKind::InputFile { file_number, variables }, span))
}

fn parse_line_input_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let variable = inner.next().unwrap().as_str().to_string();
    Ok(Statement::with_span(StmtKind::LineInput { file_number, variable }, span))
}

fn parse_select_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let expr = parse_expression(inner.next().unwrap())?;

    let mut cases = Vec::new();
    let mut default = None;

    for p in inner {
        match p.as_rule() {
            Rule::case_block => {
                let mut case_inner = p.into_inner();
                let conditions_pair = case_inner.next().unwrap();
                let mut conditions = Vec::new();

                for cond_pair in conditions_pair.into_inner() {
                    let mut cond_inner = cond_pair.into_inner();
                    let first = cond_inner.next().unwrap();

                    let condition = match first.as_rule() {
                        Rule::expression => {
                            let expr1 = parse_expression(first)?;
                            if let Some(next) = cond_inner.next() {
                                let expr2 = parse_expression(next)?;
                                CaseCondition::Range { from: expr1, to: expr2 }
                            } else {
                                CaseCondition::Value(expr1)
                            }
                        }
                        Rule::comp_op => {
                            let op = match first.as_str() {
                                "=" => ComparisonOp::Eq,
                                "<>" => ComparisonOp::NotEq,
                                "<" => ComparisonOp::Lt,
                                "<=" => ComparisonOp::LtEq,
                                ">" => ComparisonOp::Gt,
                                ">=" => ComparisonOp::GtEq,
                                _ => return Err(format!("Unknown comparison operator: {}", first.as_str())),
                            };
                            let expr = parse_expression(cond_inner.next().unwrap())?;
                            CaseCondition::Comparison { op, expr }
                        }
                        _ => return Err(format!("Unexpected rule in case condition: {:?}", first.as_rule())),
                    };
                    conditions.push(condition);
                }

                let mut body = Vec::new();
                for stmt_pair in case_inner {
                    if stmt_pair.as_rule() == Rule::statement_line {
                        for inner in stmt_pair.into_inner() {
                            if inner.as_rule() != Rule::NEWLINE && inner.as_rule() != Rule::EOI {
                                body.push(parse_statement(inner)?);
                            }
                        }
                    }
                }
                cases.push(SwitchCase { conditions, body });
            }
            Rule::case_else => {
                let mut body = Vec::new();
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::statement_line {
                        for inner in stmt_pair.into_inner() {
                            if inner.as_rule() != Rule::NEWLINE && inner.as_rule() != Rule::EOI {
                                body.push(parse_statement(inner)?);
                            }
                        }
                    }
                }
                default = Some(body);
            }
            Rule::select_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::Switch { expr, cases, default }, span))
}

// ---------------------------------------------------------------------------
// Interface / Structure / Delegate / Event parsers
// ---------------------------------------------------------------------------

fn parse_interface_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut _visibility = Visibility::Public;
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<InterfaceMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::inherits_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        parents.push(tp.as_str().to_string());
                    }
                }
            }
            Rule::interface_sub => {
                let mut sname = String::new();
                let mut params = Vec::new();
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => sname = sp.as_str().to_string(),
                        Rule::param_list => params = parse_param_list(sp)?,
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Method {
                    name: sname,
                    params,
                    return_type: None,
                    is_sub: true,
                });
            }
            Rule::interface_function => {
                let mut fname = String::new();
                let mut params = Vec::new();
                let mut ret: Option<String> = None;
                for fp in p.into_inner() {
                    match fp.as_rule() {
                        Rule::identifier => fname = fp.as_str().to_string(),
                        Rule::param_list => params = parse_param_list(fp)?,
                        Rule::type_name => ret = Some(fp.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Method {
                    name: fname,
                    params,
                    return_type: ret,
                    is_sub: false,
                });
            }
            Rule::interface_property => {
                let mut pname = String::new();
                let mut ptype: Option<String> = None;
                let mut is_readonly = false;
                let mut is_writeonly = false;
                let txt = p.as_str().to_lowercase();
                if txt.starts_with("readonly") { is_readonly = true; }
                if txt.starts_with("writeonly") { is_writeonly = true; }
                for pp in p.into_inner() {
                    match pp.as_rule() {
                        Rule::identifier => pname = pp.as_str().to_string(),
                        Rule::type_name => ptype = Some(pp.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Property {
                    name: pname,
                    type_hint: ptype,
                    is_readonly,
                    is_writeonly,
                });
            }
            Rule::interface_event => {
                let mut ename = String::new();
                let mut etype: Option<String> = None;
                for ep in p.into_inner() {
                    match ep.as_rule() {
                        Rule::identifier => ename = ep.as_str().to_string(),
                        Rule::type_name => etype = Some(ep.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Event {
                    name: ename,
                    type_hint: etype,
                });
            }
            Rule::visibility_modifier => {
                _visibility = parse_visibility(p.as_str());
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::InterfaceDecl { name, parents, members, decorators: vec![] }, span))
}

fn parse_structure_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::implements_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        interfaces.push(tp.as_str().to_string());
                    }
                }
            }
            Rule::property_decl => {
                members.push(parse_property_decl_to_member(p)?);
            }
            Rule::auto_property_decl => {
                let d = parse_auto_property_as_field(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers: Modifiers::default(),
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::sub_decl => {
                members.push(ClassMember::Method(Box::new(parse_sub_decl(p)?)));
            }
            Rule::function_decl => {
                members.push(ClassMember::Method(Box::new(parse_function_decl(p)?)));
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: Modifiers::default(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let modifiers = parse_field_modifiers(&p);
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::NEWLINE | Rule::structure_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::StructDecl { name, interfaces, members, visibility, decorators: vec![] }, span))
}

fn parse_delegate_decl(pair: Pair<Rule>, is_sub: bool) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut return_type: Option<String> = None;

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::DelegateDecl {
        name,
        params: parameters,
        return_type,
        is_sub,
        visibility,
    }, span))
}

fn parse_event_decl_to_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut event_type: Option<String> = None;

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => event_type = Some(p.as_str().to_string()),
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            _ => {}
        }
    }

    Ok(ClassMember::Event {
        name,
        type_hint: event_type,
        params: parameters,
        visibility,
    })
}

// ── Syntax Extensions Implementation ──

fn parse_synclock_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let lock_expr = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::synclock_end | Rule::NEWLINE => {}
            _ => {}
        }
    }
    Ok(Statement::with_span(StmtKind::Lock { expr: lock_expr, body }, span))
}

fn parse_query_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    // LINQ query expressions are VB-specific. Map to a Call expression with
    // a chain of method calls that the compiler can recognize.
    // For now, produce a placeholder that preserves the structure.
    let _span = to_span(&pair);
    let mut inner = pair.into_inner();
    let from_clause_pair = inner.next().unwrap();

    // Parse From clause
    let mut from_inner = from_clause_pair.into_inner();
    let mut range_var = String::new();
    let mut collection_expr = Expression::null();

    while let Some(id_pair) = from_inner.next() {
        if id_pair.as_rule() == Rule::identifier {
            range_var = id_pair.as_str().to_string();

            while let Some(p) = from_inner.next() {
                match p.as_rule() {
                    Rule::expression => { collection_expr = parse_expression(p)?; break; }
                    Rule::type_name => {} // skip
                    _ => {}
                }
            }
            break; // Take first range variable for ForIn mapping
        }
    }

    // Build the query as a series of method calls on the collection
    let query_body_pair = inner.next().unwrap();
    let body_inner = query_body_pair.into_inner();
    let mut result_expr = collection_expr.clone();

    for p in body_inner {
        match p.as_rule() {
            Rule::query_operator => {
                let op = p.into_inner().next().unwrap();
                match op.as_rule() {
                    Rule::where_clause => {
                        let filter_expr = parse_expression(op.into_inner().next().unwrap())?;
                        // .Where(Function(x) filter_expr)
                        let lambda = Expression::new(ExprKind::Lambda {
                            params: vec![Param {
                                name: range_var.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            }],
                            body: LambdaBody::Expr(Box::new(filter_expr)),
                            is_async: false,
                            captures: vec![],
                        });
                        let callee = Expression::new(ExprKind::Member {
                            object: Box::new(result_expr),
                            field: "Where".to_string(),
                            null_safe: false,
                        });
                        result_expr = Expression::new(ExprKind::Call {
                            callee: Box::new(callee),
                            args: vec![Argument::positional(lambda)],
                            optional: false,
                        });
                    }
                    Rule::order_by_clause => {
                        for ord in op.into_inner() {
                            let mut ord_inner = ord.into_inner();
                            let key_expr = parse_expression(ord_inner.next().unwrap())?;
                            let descending = matches!(
                                ord_inner.next().map(|x| x.as_str().to_lowercase()).as_deref(),
                                Some("descending")
                            );
                            let method = if descending { "OrderByDescending" } else { "OrderBy" };
                            let lambda = Expression::new(ExprKind::Lambda {
                                params: vec![Param {
                                    name: range_var.clone(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                }],
                                body: LambdaBody::Expr(Box::new(key_expr)),
                                is_async: false,
                                captures: vec![],
                            });
                            let callee = Expression::new(ExprKind::Member {
                                object: Box::new(result_expr),
                                field: method.to_string(),
                                null_safe: false,
                            });
                            result_expr = Expression::new(ExprKind::Call {
                                callee: Box::new(callee),
                                args: vec![Argument::positional(lambda)],
                                optional: false,
                            });
                        }
                    }
                    Rule::let_clause => {
                        let mut let_inner = op.into_inner();
                        let _let_name = let_inner.next().unwrap().as_str().to_string();
                        let _let_value = parse_expression(let_inner.next().unwrap())?;
                        // Let clause is more complex to map; keep result_expr as-is for now
                    }
                    _ => {}
                }
            }
            Rule::select_or_group_clause => {
                let inner_sg = p.into_inner().next().unwrap();
                match inner_sg.as_rule() {
                    Rule::select_clause => {
                        let exprs: Vec<Expression> = inner_sg.into_inner()
                            .map(|x| parse_expression(x))
                            .collect::<Result<Vec<_>,_>>()?;
                        if !exprs.is_empty() {
                            // .Select(Function(x) expr)
                            let select_body = if exprs.len() == 1 {
                                exprs.into_iter().next().unwrap()
                            } else {
                                // Multiple select expressions → tuple-like
                                Expression::new(ExprKind::Array(
                                    exprs.into_iter().map(|e| ArrayElement {
                                        key: None, value: e, spread: false, by_ref: false,
                                    }).collect()
                                ))
                            };
                            let lambda = Expression::new(ExprKind::Lambda {
                                params: vec![Param {
                                    name: range_var.clone(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                }],
                                body: LambdaBody::Expr(Box::new(select_body)),
                                is_async: false,
                                captures: vec![],
                            });
                            let callee = Expression::new(ExprKind::Member {
                                object: Box::new(result_expr),
                                field: "Select".to_string(),
                                null_safe: false,
                            });
                            result_expr = Expression::new(ExprKind::Call {
                                callee: Box::new(callee),
                                args: vec![Argument::positional(lambda)],
                                optional: false,
                            });
                        }
                    }
                    Rule::group_clause => {
                        let mut exprs = Vec::new();
                        for x in inner_sg.into_inner() {
                            if x.as_rule() == Rule::expression {
                                exprs.push(parse_expression(x)?);
                            }
                        }
                        if exprs.len() >= 2 {
                            let key = exprs.pop().unwrap();
                            let _item = exprs.pop().unwrap();
                            // .GroupBy(Function(x) key)
                            let lambda = Expression::new(ExprKind::Lambda {
                                params: vec![Param {
                                    name: range_var.clone(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                }],
                                body: LambdaBody::Expr(Box::new(key)),
                                is_async: false,
                                captures: vec![],
                            });
                            let callee = Expression::new(ExprKind::Member {
                                object: Box::new(result_expr),
                                field: "GroupBy".to_string(),
                                null_safe: false,
                            });
                            result_expr = Expression::new(ExprKind::Call {
                                callee: Box::new(callee),
                                args: vec![Argument::positional(lambda)],
                                optional: false,
                            });
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(result_expr)
}

fn parse_xml_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    // XML literals are VB-specific. Convert to a string representation.
    let span = to_span(&pair);
    let xml_text = pair.as_str().to_string();
    Ok(Expression::with_span(ExprKind::Lit(Literal::Str(xml_text)), span))
}

fn parse_l_value_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let source = pair.as_str().trim();
    let bytes = source.as_bytes();
    let mut cursor = 0usize;

    while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    let start = cursor;
    while cursor < bytes.len()
        && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'_')
    {
        cursor += 1;
    }

    if start == cursor {
        return Err("l_value_expression missing identifier".to_string());
    }

    let mut expr = Expression::ident(&source[start..cursor]);

    while cursor < bytes.len() {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            break;
        }

        match bytes[cursor] {
            b'.' => {
                cursor += 1;
                while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
                    cursor += 1;
                }
                let name_start = cursor;
                while cursor < bytes.len()
                    && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'_')
                {
                    cursor += 1;
                }
                if name_start == cursor {
                    return Err("l_value_expression missing member name".to_string());
                }
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: source[name_start..cursor].to_string(),
                    null_safe: false,
                });
            }
            b'(' => {
                let args_start = cursor + 1;
                let mut depth = 1usize;
                cursor += 1;
                while cursor < bytes.len() && depth > 0 {
                    match bytes[cursor] {
                        b'(' => depth += 1,
                        b')' => depth -= 1,
                        _ => {}
                    }
                    cursor += 1;
                }
                if depth != 0 {
                    return Err("l_value_expression missing closing ')'".to_string());
                }

                let args_text = source[args_start..cursor - 1].trim();
                let mut args = Vec::new();
                if !args_text.is_empty() {
                    let mut parsed = VbParser::parse(Rule::argument_list, args_text)
                        .map_err(|err| err.to_string())?;
                    let arg_list_pair = parsed
                        .next()
                        .ok_or_else(|| "l_value_expression missing argument list".to_string())?;
                    args = parse_argument_list(arg_list_pair)?
                        .into_iter()
                        .map(|arg| arg.value)
                        .collect();
                }

                if let ExprKind::Member { object, field, .. } = &expr.kind {
                    if field.eq_ignore_ascii_case("Item") && !args.is_empty() {
                        let mut indexed = (**object).clone();
                        for idx_expr in args {
                            indexed = Expression::new(ExprKind::Index {
                                object: Box::new(indexed),
                                index: Box::new(idx_expr),
                                null_safe: false,
                            });
                        }
                        expr = indexed;
                        continue;
                    }
                }

                if args.len() == 1 {
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(args.into_iter().next().unwrap()),
                        null_safe: false,
                    });
                } else if !args.is_empty() {
                    for idx_expr in args {
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(idx_expr),
                            null_safe: false,
                        });
                    }
                } else {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: vec![],
                        optional: false,
                    });
                }
            }
            _ => break,
        }
    }

    Ok(expr)
}

// ── Helper functions ──

/// Split a `ctrl.Event` (or `obj.Sub.Event`) string into a control expression
/// and a lowercase event name. The last segment is the event; everything
/// before becomes the control expression (member chain).
fn split_event_target(s: &str) -> (Expression, String) {
    let parts: Vec<&str> = s.split('.').collect();
    if parts.len() < 2 {
        return (Expression::ident(s), String::new());
    }
    let event = parts[parts.len() - 1].to_lowercase();
    let control = build_dotted_expr(&parts[..parts.len() - 1].join("."));
    (control, event)
}

/// Build an Expression from a dotted name like `me.btn1` or `obj.field.method`.
/// The first segment becomes an Ident; subsequent segments become Member access.
fn build_dotted_expr(s: &str) -> Expression {
    let parts: Vec<&str> = s.split('.').collect();
    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or("");
    let mut expr = Expression::ident(first);
    for seg in iter {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: seg.to_string(),
            null_safe: false,
        });
    }
    expr
}

fn parse_visibility(s: &str) -> Visibility {
    match s.to_lowercase().as_str() {
        "public" => Visibility::Public,
        "private" => Visibility::Private,
        "protected" => Visibility::Protected,
        "friend" => Visibility::Internal,
        _ => Visibility::Public,
    }
}

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
