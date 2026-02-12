use pest::Parser;
use pest::iterators::Pair;
use pest_derive::Parser;
use crate::ast::*;

#[derive(Parser)]
#[grammar = "grammar.pest"]
pub struct VBParser;

#[derive(Debug, thiserror::Error)]
pub enum ParseError {
    #[error("Pest parsing error: {0}")]
    PestError(#[from] pest::error::Error<Rule>),

    #[error("Unexpected rule: {0:?}")]
    UnexpectedRule(Rule),

    #[error("Parse error: {0}")]
    Custom(String),
}

pub type ParseResult<T> = Result<T, ParseError>;

pub fn parse_program(source: &str) -> ParseResult<Program> {
    // Strip BOM from any source — single place for all callers
    let source = source.trim_start_matches('\u{feff}');
    let pairs = VBParser::parse(Rule::program, source)?;
    let mut declarations = Vec::new();
    let mut statements = Vec::new();

    for pair in pairs {
        match pair.as_rule() {
            Rule::program => {
                for inner in pair.into_inner() {
                    match inner.as_rule() {
                        Rule::statement_line => {
                            for stmt_pair in inner.into_inner() {
                                if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                                    continue;
                                }
                                if stmt_pair.as_rule() == Rule::module_decl {
                                    // Flatten module contents into top-level declarations
                                    declarations.extend(parse_module_decl(stmt_pair)?);
                                } else if let Some(decl) = try_parse_declaration(stmt_pair.clone())? {
                                    declarations.push(decl);
                                } else {
                                    statements.push(parse_statement(stmt_pair)?);
                                }
                            }
                        }
                        Rule::NEWLINE | Rule::EOI => {}
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Program {
        declarations,
        statements,
    })
}

pub fn parse_expression_str(source: &str) -> ParseResult<Expression> {
    let mut pairs = VBParser::parse(Rule::expression, source)?;
    let pair = pairs.next().ok_or_else(|| ParseError::Custom("No expression found".to_string()))?;
    parse_expression(pair)
}

fn try_parse_declaration(pair: Pair<Rule>) -> ParseResult<Option<Declaration>> {
    match pair.as_rule() {
        Rule::dim_statement => {
            Ok(Some(Declaration::Variable(parse_dim_statement(pair)?)))
        }
        Rule::const_statement => {
            Ok(Some(Declaration::Constant(parse_const_statement(pair)?)))
        }
        Rule::sub_decl => Ok(Some(Declaration::Sub(parse_sub_decl(pair)?))),
        Rule::function_decl => Ok(Some(Declaration::Function(parse_function_decl(pair)?))),
        Rule::class_decl => Ok(Some(Declaration::Class(parse_class_decl(pair)?))),
        Rule::enum_decl => Ok(Some(Declaration::Enum(parse_enum_decl(pair)?))),
        _ => Ok(None),
    }
}

fn parse_dim_statement(pair: Pair<Rule>) -> ParseResult<VariableDecl> {
    let inner = pair.into_inner();
    let mut name = Identifier::new("");
    let mut var_type = None;
    let mut array_bounds = None;
    let mut initializer = None;
    let mut is_new = false;
    let mut ctor_args: Vec<Expression> = Vec::new();
    let mut from_init: Option<Vec<Expression>> = None;
    let mut with_init: Option<Vec<(String, Expression)>> = None;

    for p in inner {
        match p.as_rule() {
            Rule::dim_new_keyword => {
                is_new = true;
                continue;
            }
            Rule::identifier => {
                name = Identifier::new(p.as_str());
            }
            Rule::array_bounds => {
                let bounds: Vec<Expression> = p.into_inner()
                    .map(|bound_expr| parse_expression(bound_expr))
                    .collect::<Result<_, _>>()?;
                array_bounds = Some(bounds);
            }
            Rule::type_name => {
                var_type = Some(VBType::from_str(p.as_str()));
            }
            Rule::argument_list => {
                // Constructor arguments for "As New Type(args)"
                for arg_pair in p.into_inner() {
                    if arg_pair.as_rule() == Rule::expression {
                        ctor_args.push(parse_expression(arg_pair)?);
                    }
                }
            }
            Rule::from_initializer => {
                let elements: Vec<Expression> = p.into_inner()
                    .filter(|e| e.as_rule() == Rule::expression)
                    .map(|e| parse_expression(e))
                    .collect::<Result<Vec<_>, _>>()?;
                from_init = Some(elements);
            }
            Rule::with_initializer => {
                let mut members = Vec::new();
                for mi in p.into_inner() {
                    if mi.as_rule() != Rule::member_initializer { continue; }
                    let mut mi_inner = mi.into_inner();
                    let prop_name = mi_inner.next().unwrap().as_str().to_string();
                    let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                    members.push((prop_name, prop_expr));
                }
                with_init = Some(members);
            }
            Rule::array_literal => {
                initializer = Some(parse_array_literal(p)?);
            }
            Rule::expression => {
                initializer = Some(parse_expression(p)?);
            }
            _ => {}
        }
    }

    // Handle "As New Type" syntax
    if is_new && initializer.is_none() {
        if let Some(t) = &var_type {
            let class_id = Identifier::new(t.to_string());
            if let Some(elements) = from_init {
                initializer = Some(Expression::NewFromInitializer(class_id, ctor_args, elements));
            } else if let Some(members) = with_init {
                initializer = Some(Expression::NewWithInitializer(class_id, ctor_args, members));
            } else {
                initializer = Some(Expression::New(class_id, ctor_args));
            }
        }
    }

    Ok(VariableDecl {
        name,
        var_type,
        array_bounds,
        initializer,
    })
}

fn parse_const_statement(pair: Pair<Rule>) -> ParseResult<ConstDecl> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = Identifier::new("");
    let mut const_type = VBType::Variant;
    let mut value = None;

    for p in inner {
        match p.as_rule() {
            Rule::visibility_modifier => {
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Friend,
                    _ => {}
                }
            }
            Rule::identifier => {
                name = Identifier::new(p.as_str());
            }
            Rule::type_name => {
                const_type = VBType::from_str(p.as_str());
            }
            Rule::expression => {
                value = Some(parse_expression(p)?);
            }
            _ => {}
        }
    }

    Ok(ConstDecl {
        visibility,
        name,
        const_type,
        value: value.ok_or_else(|| ParseError::Custom("Const must have a value".to_string()))?,
    })
}

fn parse_array_literal(pair: Pair<Rule>) -> ParseResult<Expression> {
    let elements: Vec<Expression> = pair.into_inner()
        .map(|p| parse_expression(p))
        .collect::<Result<_, _>>()?;
    Ok(Expression::ArrayLiteral(elements))
}

fn parse_redim_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let inner = pair.into_inner();
    let mut preserve = false;
    let mut array = Identifier::new("");
    let mut bounds = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::preserve_keyword => {
                preserve = true;
            }
            Rule::identifier => {
                array = Identifier::new(p.as_str());
            }
            Rule::array_bounds => {
                bounds = p.into_inner()
                    .map(|bound_expr| parse_expression(bound_expr))
                    .collect::<Result<_, _>>()?;
            }
            _ => {}
        }
    }

    Ok(Statement::ReDim {
        preserve,
        array,
        bounds,
    })
}




fn parse_sub_decl(pair: Pair<Rule>) -> ParseResult<SubDecl> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = Identifier::new("");
    let mut parameters = Vec::new();
    let mut body = Vec::new();
    let mut handles: Option<Vec<String>> = None;
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
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Friend,
                    _ => {}
                }
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
            }
            Rule::sub_name => name = Identifier::new(p.as_str()),
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
                        Rule::statement_line => {
                            for inner in stmt_pair.into_inner() {
                                if inner.as_rule() == Rule::NEWLINE || inner.as_rule() == Rule::EOI {
                                    continue;
                                }
                                body.push(parse_statement(inner)?);
                            }
                        }
                        Rule::sub_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => {
                            body.push(parse_statement(stmt_pair)?);
                        }
                    }
                }
            }
            Rule::handles_clause => {
                let mut handle_list = Vec::new();
                for hp in p.into_inner() {
                    if hp.as_rule() == Rule::dotted_identifier {
                        handle_list.push(hp.as_str().to_string());
                    }
                }
                if !handle_list.is_empty() {
                    handles = Some(handle_list);
                }
            }
            _ => {}
        }
    }

    Ok(SubDecl {
        visibility,
        name,
        parameters,
        body,
        handles,
        is_async,
        is_extension,
        is_overridable,
        is_overrides,
        is_must_override,
        is_shared,
        is_not_overridable,
    })
}

fn parse_function_decl(pair: Pair<Rule>) -> ParseResult<FunctionDecl> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = Identifier::new("");
    let mut parameters = Vec::new();
    let mut return_type = None;
    let mut body = Vec::new();
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
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Friend,
                    _ => {}
                }
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
            Rule::identifier => name = Identifier::new(p.as_str()),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => return_type = Some(VBType::from_str(p.as_str())),
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
                        Rule::statement_line => {
                            for inner in stmt_pair.into_inner() {
                                if inner.as_rule() == Rule::NEWLINE || inner.as_rule() == Rule::EOI {
                                    continue;
                                }
                                body.push(parse_statement(inner)?);
                            }
                        }
                        Rule::func_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => {
                            body.push(parse_statement(stmt_pair)?);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(FunctionDecl {
        visibility,
        name,
        parameters,
        return_type,
        body,
        is_async,
        is_extension,
        is_overridable,
        is_overrides,
        is_must_override,
        is_shared,
        is_not_overridable,
    })
}

fn parse_module_decl(pair: Pair<Rule>) -> ParseResult<Vec<Declaration>> {
    let mut declarations = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::sub_decl => declarations.push(Declaration::Sub(parse_sub_decl(p)?)),
            Rule::function_decl => declarations.push(Declaration::Function(parse_function_decl(p)?)),
            Rule::const_statement => declarations.push(Declaration::Constant(parse_const_statement(p)?)),
            Rule::dim_statement => declarations.push(Declaration::Variable(parse_dim_statement(p)?)),
            Rule::field_decl => declarations.push(Declaration::Variable(parse_field_decl(p)?)),
            Rule::class_decl => declarations.push(Declaration::Class(parse_class_decl(p)?)),
            Rule::enum_decl => declarations.push(Declaration::Enum(parse_enum_decl(p)?)),
            Rule::identifier | Rule::NEWLINE | Rule::module_end => {}
            _ => {}
        }
    }

    Ok(declarations)
}

fn parse_class_decl(pair: Pair<Rule>) -> ParseResult<ClassDecl> {
    let inner = pair.into_inner();
    let mut name = Identifier::new("");
    let mut is_partial = false;
    let mut visibility = Visibility::Public;
    let mut inherits = None;
    let mut implements = Vec::new();
    let mut properties = Vec::new();
    let mut methods = Vec::new();
    let mut fields = Vec::new();
    let mut is_must_inherit = false;
    let mut is_not_inheritable = false;

    for p in inner {
        match p.as_rule() {
            Rule::partial_keyword => is_partial = true,
            Rule::visibility_modifier => {
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Friend,
                    _ => {}
                }
            }
            Rule::must_inherit_keyword => is_must_inherit = true,
            Rule::not_inheritable_keyword => is_not_inheritable = true,
            Rule::inherits_statement => {
                 if let Some(type_pair) = p.into_inner().next() {
                    inherits = Some(VBType::from_str(type_pair.as_str()));
                }
            }
            Rule::implements_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        implements.push(VBType::from_str(tp.as_str()));
                    }
                }
            }
            Rule::identifier => name = Identifier::new(p.as_str()),
            Rule::property_decl => properties.push(parse_property_decl(p)?),
            Rule::sub_decl => methods.push(MethodDecl::Sub(parse_sub_decl(p)?)),
            Rule::function_decl => methods.push(MethodDecl::Function(parse_function_decl(p)?)),
            Rule::dim_statement => {
                if let Ok(crate::ast::stmt::Statement::Dim(decl)) = parse_statement(p) {
                    fields.push(decl);
                }
            }
            Rule::field_decl => {
                fields.push(parse_field_decl(p)?);
            }
            Rule::NEWLINE | Rule::class_end => {}
            _ => {}
        }
    }

    Ok(ClassDecl {
        visibility,
        name,
        is_partial,
        inherits,
        implements,
        properties,
        methods,
        fields,
        is_must_inherit,
        is_not_inheritable,
    })
}

fn parse_property_decl(pair: Pair<Rule>) -> ParseResult<PropertyDecl> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = Identifier::new("");
    let mut parameters = Vec::new();
    let mut return_type = None;
    let mut getter = None;
    let mut setter = None;

    for p in inner {
        match p.as_str().to_lowercase().as_str() {
            "public" => visibility = Visibility::Public,
            "private" => visibility = Visibility::Private,
            _ => {
                match p.as_rule() {
                    Rule::identifier => name = Identifier::new(p.as_str()),
                    Rule::param_list => parameters = parse_param_list(p)?,
                    Rule::type_name => return_type = Some(VBType::from_str(p.as_str())),
                    Rule::property_get => getter = Some(parse_property_get(p)?),
                    Rule::property_set => setter = Some(parse_property_set(p)?),
                    _ => {}
                }
            }
        }
    }

    Ok(PropertyDecl {
        visibility,
        name,
        parameters,
        return_type,
        getter,
        setter,
    })
}

fn parse_property_get(pair: Pair<Rule>) -> ParseResult<Vec<Statement>> {
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

fn parse_property_set(pair: Pair<Rule>) -> ParseResult<(Parameter, Vec<Statement>)> {
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
    
    Ok((param, body))
}

fn parse_param_list(pair: Pair<Rule>) -> ParseResult<Vec<Parameter>> {
    pair.into_inner().map(parse_parameter).collect()
}

fn parse_parameter(pair: Pair<Rule>) -> ParseResult<Parameter> {
    let inner = pair.into_inner();
    let mut pass_type = ParameterPassType::ByRef;
    let mut name = Identifier::new("");
    let mut param_type = None;
    let mut is_optional = false;
    let mut default_value = None;
    let mut is_nullable = false;

    for p in inner {
        match p.as_rule() {
            Rule::pass_type_keyword => {
                let text = p.as_str().to_lowercase();
                if text == "byval" {
                    pass_type = ParameterPassType::ByVal;
                } else {
                    pass_type = ParameterPassType::ByRef;
                }
            }
            Rule::optional_keyword => {
                is_optional = true;
            }
            Rule::identifier => {
                name = Identifier::new(p.as_str());
            }
            Rule::type_name => param_type = Some(VBType::from_str(p.as_str())),
            Rule::nullable_marker => is_nullable = true,
            Rule::expression => default_value = Some(parse_expression(p)?),
            _ => {}
        }
    }

    Ok(Parameter {
        pass_type,
        name,
        param_type,
        is_optional,
        default_value,
        is_nullable,
    })
}

fn parse_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    match pair.as_rule() {
        Rule::dim_statement => {
            Ok(Statement::Dim(parse_dim_statement(pair)?))
        }
        Rule::const_statement => {
            Ok(Statement::Const(parse_const_statement(pair)?))
        }
        Rule::redim_statement => {
            parse_redim_statement(pair)
        }
        Rule::select_statement => {
            parse_select_statement(pair)
        }
        Rule::dot_assign_statement => {
            // .prop1.prop2 = value (inside With block)
            let inner = pair.into_inner();
            let mut members = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(Identifier::new(p.as_str())),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| ParseError::Custom("dot_assign missing value".to_string()))?;
            if members.is_empty() {
                return Err(ParseError::Custom("dot_assign needs at least one member".to_string()));
            }
            let last = members.pop().unwrap();
            let mut obj = Expression::WithTarget;
            for m in members {
                obj = Expression::MemberAccess(Box::new(obj), m);
            }
            Ok(Statement::MemberAssignment {
                object: obj,
                member: last,
                value,
            })
        }
        Rule::me_assign_statement => {
            // Me.prop1.prop2 = value
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut members = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(Identifier::new(p.as_str())),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| ParseError::Custom("me_assign_statement missing value".to_string()))?;
            if members.is_empty() {
                return Err(ParseError::Custom("me_assign_statement needs at least one member".to_string()));
            }
            let last = members.pop().unwrap();
            let mut obj: Expression = Expression::Me;
            for m in members {
                obj = Expression::MemberAccess(Box::new(obj), m);
            }
            Ok(Statement::MemberAssignment {
                object: obj,
                member: last,
                value,
            })
        }
        Rule::mybase_assign_statement => {
            // MyBase.prop = value
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut members = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => members.push(Identifier::new(p.as_str())),
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| ParseError::Custom("mybase_assign_statement missing value".to_string()))?;
            if members.is_empty() {
                return Err(ParseError::Custom("mybase_assign_statement needs at least one member".to_string()));
            }
            let last = members.pop().unwrap();
            let mut obj: Expression = Expression::MyBase;
            for m in members {
                obj = Expression::MemberAccess(Box::new(obj), m);
            }
            Ok(Statement::MemberAssignment {
                object: obj,
                member: last,
                value,
            })
        }
        Rule::assign_statement => {
            let mut inner = pair.into_inner();
            let target = Identifier::new(inner.next().unwrap().as_str());

            // Collect all parts: identifiers (for member access) and expressions (for array indices or value)
            let mut members = Vec::new();
            let mut indices = Vec::new();
            let mut expressions = Vec::new();

            for p in inner {
                match p.as_rule() {
                    Rule::identifier => {
                        members.push(Identifier::new(p.as_str()));
                    }
                    Rule::expression => {
                        expressions.push(parse_expression(p)?);
                    }
                    _ => {}
                }
            }

            // The last expression is always the value being assigned
            let value = expressions.pop().unwrap();

            // If there are remaining expressions, they are array indices
            if !expressions.is_empty() {
                indices = expressions;
            }

            // Determine the type of assignment
            if !indices.is_empty() {
                // Array assignment: arr(5) = value
                Ok(Statement::ArrayAssignment {
                    array: target,
                    indices,
                    value,
                })
            } else if members.is_empty() {
                // Simple assignment: x = value
                Ok(Statement::Assignment {
                    target,
                    value,
                })
            } else {
                // Member assignment: obj.prop = value
                let mut obj = Expression::Variable(target);
                for i in 0..members.len() - 1 {
                    obj = Expression::MemberAccess(Box::new(obj), members[i].clone());
                }
                Ok(Statement::MemberAssignment {
                    object: obj,
                    member: members.last().unwrap().clone(),
                    value,
                })
            }
        }
        Rule::set_statement => {
            let mut inner = pair.into_inner();
            let target = Identifier::new(inner.next().unwrap().as_str());
            let value = parse_expression(inner.next().unwrap())?;

            Ok(Statement::SetAssignment { target, value })
        }
        Rule::compound_assign_statement => {
            let mut inner = pair.into_inner();
            let target = Identifier::new(inner.next().unwrap().as_str());

            let mut members = Vec::new();
            let mut indices = Vec::new();
            let mut op = CompoundOp::AddAssign;
            let mut expressions = Vec::new();

            for p in inner {
                match p.as_rule() {
                    Rule::identifier => members.push(Identifier::new(p.as_str())),
                    Rule::compound_assign_op => {
                        op = match p.as_str() {
                            "+=" => CompoundOp::AddAssign,
                            "-=" => CompoundOp::SubtractAssign,
                            "*=" => CompoundOp::MultiplyAssign,
                            "/=" => CompoundOp::DivideAssign,
                            "\\=" => CompoundOp::IntDivideAssign,
                            "&=" => CompoundOp::ConcatAssign,
                            "^=" => CompoundOp::ExponentAssign,
                            "<<=" => CompoundOp::ShiftLeftAssign,
                            ">>=" => CompoundOp::ShiftRightAssign,
                            _ => CompoundOp::AddAssign,
                        };
                    }
                    Rule::expression => expressions.push(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = expressions.pop().unwrap();
            if !expressions.is_empty() {
                indices = expressions;
            }

            Ok(Statement::CompoundAssignment {
                target,
                members,
                indices,
                operator: op,
                value,
            })
        }
        Rule::raiseevent_statement => {
            let mut inner = pair.into_inner();
            let event_name = Identifier::new(inner.next().unwrap().as_str());
            let mut arguments = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for arg in p.into_inner() {
                        arguments.push(parse_expression(arg)?);
                    }
                }
            }
            Ok(Statement::RaiseEvent { event_name, arguments })
        }
        Rule::if_statement => parse_if_statement(pair),
        Rule::single_line_if_statement => parse_single_line_if(pair),
        Rule::for_each_statement => parse_for_each_statement(pair),
        Rule::for_statement => parse_for_statement(pair),
        Rule::while_statement => parse_while_statement(pair),
        Rule::do_loop_statement => parse_do_loop_statement(pair),
        Rule::with_statement => parse_with_statement(pair),
        Rule::using_statement => parse_using_statement(pair),
        Rule::exit_statement => {
            let mut inner = pair.into_inner();
            let exit_type = inner.next()
                .ok_or_else(|| ParseError::Custom("Exit statement missing type".to_string()))?
                .as_str()
                .to_lowercase();
                
            match exit_type.as_str() {
                "sub" => Ok(Statement::ExitSub),
                "function" => Ok(Statement::ExitFunction),
                "for" => Ok(Statement::ExitFor),
                "do" => Ok(Statement::ExitDo),
                "while" => Ok(Statement::ExitWhile),
                "select" => Ok(Statement::ExitSelect),
                "try" => Ok(Statement::ExitTry),
                "property" => Ok(Statement::ExitProperty),
                _ => Err(ParseError::Custom(format!("Unknown exit type: {}", exit_type))),
            }
        }
        Rule::try_statement => parse_try_statement(pair),
        Rule::throw_statement => {
            let mut inner = pair.into_inner();
            let expr = inner.next().map(parse_expression).transpose()?;
            Ok(Statement::Throw(expr))
        }
        Rule::continue_statement => parse_continue_statement(pair),
        Rule::open_statement => parse_open_statement(pair),
        Rule::close_statement => parse_close_statement(pair),
        Rule::print_file_statement => parse_print_file_statement(pair),
        Rule::write_file_statement => parse_write_file_statement(pair),
        Rule::input_file_statement => parse_input_file_statement(pair),
        Rule::line_input_statement => parse_line_input_statement(pair),
        Rule::return_statement => {
            let mut inner = pair.into_inner();
            let value = inner.next().map(parse_expression).transpose()?;
            Ok(Statement::Return(value))
        }
        Rule::call_statement => {
            let mut inner = pair.into_inner();
            let first = inner.next().unwrap();

            // Check if it's a member_call, member_access, call_expression, me_member_call, cast_member_call, or simple identifier
            match first.as_rule() {
                Rule::cast_member_call | Rule::me_member_call | Rule::mybase_member_call | Rule::member_call | Rule::member_access | Rule::call_expression => {
                    // Parse as expression and convert to statement
                    let expr = parse_expression(first)?;
                    Ok(Statement::ExpressionStatement(expr))
                }
                Rule::identifier => {
                    // Could be: identifier, identifier(args), or identifier args
                    let name = Identifier::new(first.as_str());
                    let arguments = inner.next()
                        .map(|p| {
                            if p.as_rule() == Rule::argument_list {
                                parse_argument_list(p)
                            } else {
                                // Single expression argument
                                parse_expression(p).map(|e| vec![e])
                            }
                        })
                        .transpose()?
                        .unwrap_or_default();

                    Ok(Statement::Call { name, arguments })
                }
                _ => {
                    let name = Identifier::new(first.as_str());
                    Ok(Statement::Call { name, arguments: vec![] })
                }
            }
        }
        Rule::expression_statement => {
            let expr = parse_expression(pair.into_inner().next().unwrap())?;
            Ok(Statement::ExpressionStatement(expr))
        }
        Rule::addhandler_statement => {
            let mut inner = pair.into_inner();
            let event_target = inner.next().unwrap().as_str().to_string();
            let addressof = inner.next().unwrap(); // addressof_expr
            let handler = addressof.into_inner().next().unwrap().as_str().to_string();
            Ok(Statement::AddHandler { event_target, handler })
        }
        Rule::removehandler_statement => {
            let mut inner = pair.into_inner();
            let event_target = inner.next().unwrap().as_str().to_string();
            let addressof = inner.next().unwrap(); // addressof_expr
            let handler = addressof.into_inner().next().unwrap().as_str().to_string();
            Ok(Statement::RemoveHandler { event_target, handler })
        }
        // New declarations — parse gracefully as no-op statements for now
        Rule::interface_decl | Rule::structure_decl | Rule::namespace_decl |
        Rule::event_decl | Rule::delegate_sub_decl | Rule::delegate_function_decl => {
            // These are parsed by the grammar but the runtime doesn't execute them yet.
            // Return an expression statement with Nothing to avoid breaking parsing.
            Ok(Statement::ExpressionStatement(Expression::Nothing))
        }
        _ => Err(ParseError::UnexpectedRule(pair.as_rule())),
    }
}

fn parse_if_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let condition = parse_expression(inner.next().unwrap())?;
    let mut then_branch = Vec::new();
    let mut elseif_branches = Vec::new();
    let mut else_branch = None;

    for p in inner {
        match p.as_rule() {
            Rule::if_body => {
                if then_branch.is_empty() {
                    then_branch = parse_block(p)?;
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
                    elseif_branches.push((cond, elseif_body));
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
                else_branch = Some(body);
            }
            Rule::NEWLINE | Rule::if_end => {}
            _ => {}
        }
    }

    Ok(Statement::If {
        condition,
        then_branch,
        elseif_branches,
        else_branch,
    })
}

fn parse_block(pair: Pair<Rule>) -> ParseResult<Vec<Statement>> {
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

fn parse_for_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let variable = Identifier::new(inner.next().unwrap().as_str());

    // Skip optional 'As type_name'
    let mut next = inner.next().unwrap();
    if next.as_rule() == Rule::type_name {
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
            _ => {}
        }
    }

    Ok(Statement::For {
        variable,
        start,
        end,
        step,
        body,
    })
}

fn parse_while_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let condition = parse_expression(inner.next().unwrap())?;
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

    Ok(Statement::While { condition, body })
}

fn parse_do_loop_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let inner = pair.into_inner();
    let mut pre_condition = None;
    let mut post_condition = None;
    let mut body = Vec::new();
    let mut current_loop_type = LoopConditionType::While;

    for p in inner {
        match p.as_rule() {
            Rule::do_while_kw => current_loop_type = LoopConditionType::While,
            Rule::do_until_kw => current_loop_type = LoopConditionType::Until,
            Rule::expression => {
                // Determine if it's pre or post condition based on position
                if body.is_empty() {
                    pre_condition = Some((current_loop_type, parse_expression(p)?));
                } else {
                    post_condition = Some((current_loop_type, parse_expression(p)?));
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
                        Rule::do_while_kw => current_loop_type = LoopConditionType::While,
                        Rule::do_until_kw => current_loop_type = LoopConditionType::Until,
                        Rule::expression => {
                            post_condition = Some((current_loop_type, parse_expression(dp)?));
                        }
                        _ => {}
                    }
                }
            }
            Rule::NEWLINE => {}
            _ => {}
        }
    }

    Ok(Statement::DoLoop {
        pre_condition,
        body,
        post_condition,
    })
}

fn parse_expression(pair: Pair<Rule>) -> ParseResult<Expression> {
    match pair.as_rule() {
        Rule::expression | Rule::logical_xor | Rule::logical_or | Rule::logical_and |
        Rule::equality | Rule::comparison | Rule::bit_shift | Rule::additive |
        Rule::multiplicative | Rule::exponent => {
            parse_binary_expression(pair)
        }
        Rule::not_condition => {
            // not_condition = { not_op? ~ equality | equality }
            let mut inner = pair.into_inner();
            let first = inner.next().unwrap();
            if first.as_rule() == Rule::not_op {
                let operand = parse_expression(inner.next().unwrap())?;
                Ok(Expression::Not(Box::new(operand)))
            } else {
                parse_expression(first)
            }
        }
        Rule::lambda_expression => {
            parse_lambda_expression(pair)
        }
        Rule::typeof_expression => {
            let mut inner = pair.into_inner();
            let expr = parse_expression(inner.next().unwrap())?;
            let type_name = inner.next().unwrap().as_str().trim().to_string();
            Ok(Expression::TypeOf {
                expr: Box::new(expr),
                type_name,
            })
        }

        Rule::unary => {
            parse_unary_expression(pair)
        }
        Rule::postfix => {
            parse_postfix_expression(pair)
        }
        Rule::call_expression => {
            let mut inner = pair.into_inner();
            let name = Identifier::new(inner.next().unwrap().as_str());
            let arguments = inner.next()
                .map(parse_argument_list)
                .transpose()?
                .unwrap_or_default();

            Ok(Expression::Call(name, arguments))
        }
        Rule::member_call => {
            let mut inner = pair.into_inner();
            // First child is always the root identifier
            let first = inner.next().unwrap();
            let mut expr = Expression::Variable(Identifier::new(first.as_str()));

            // Remaining children are member_chain segments
            for chain in inner {
                expr = parse_member_chain_node(chain, expr)?;
            }

            Ok(expr)
        }
        Rule::member_access => {
            let mut inner = pair.into_inner();
            let mut expr = Expression::Variable(Identifier::new(inner.next().unwrap().as_str()));

            for p in inner {
                expr = Expression::MemberAccess(Box::new(expr), Identifier::new(p.as_str()));
            }

            Ok(expr)
        }
        Rule::identifier => Ok(Expression::Variable(Identifier::new(pair.as_str()))),
        Rule::cast_expression => {
            let text = pair.as_str();
            let kind = if text[..5].eq_ignore_ascii_case("CType") {
                crate::ast::expr::CastKind::CType
            } else if text[..10].eq_ignore_ascii_case("DirectCast") {
                crate::ast::expr::CastKind::DirectCast
            } else {
                crate::ast::expr::CastKind::TryCast
            };
            let mut inner = pair.into_inner();
            let expr = parse_expression(inner.next().unwrap())?;
            let type_name = inner.next().unwrap().as_str().to_string();
            Ok(Expression::Cast {
                kind,
                expr: Box::new(expr),
                target_type: type_name,
            })
        }
        Rule::cast_member_call => {
            let mut inner = pair.into_inner();
            let cast_pair = inner.next().unwrap();
            let mut expr = parse_expression(cast_pair)?;
            // Remaining children are member_chain segments
            for chain in inner {
                expr = parse_member_chain_node(chain, expr)?;
            }
            Ok(expr)
        }
        Rule::interpolated_string => {
            // $"text {expr} text {expr} text" → concatenation
            let s = pair.as_str();
            // Strip $" prefix and " suffix
            let inner = &s[2..s.len()-1];
            // Replace VB doubled quotes
            let inner = inner.replace("\"\"", "\"");
            
            // Split into text parts and {expression} parts
            let mut parts: Vec<Expression> = Vec::new();
            let mut current_text = String::new();
            let mut chars = inner.chars().peekable();
            
            while let Some(ch) = chars.next() {
                if ch == '{' {
                    // Check for {{ escape (literal brace)
                    if chars.peek() == Some(&'{') {
                        chars.next();
                        current_text.push('{');
                        continue;
                    }
                    // Flush text so far
                    if !current_text.is_empty() {
                        parts.push(Expression::StringLiteral(current_text.clone()));
                        current_text.clear();
                    }
                    // Collect expression until matching }
                    let mut expr_text = String::new();
                    let mut depth = 1;
                    while let Some(c) = chars.next() {
                        if c == '{' { depth += 1; }
                        if c == '}' { depth -= 1; if depth == 0 { break; } }
                        expr_text.push(c);
                    }
                    // Parse the expression text as a VB expression
                    let expr_code = format!("Sub _Tmp()\nDim _x = {}\nEnd Sub", expr_text);
                    match crate::parse_program(&expr_code) {
                        Ok(program) => {
                            // Extract the expression from Dim _x = <expr>
                            let mut found = false;
                            for decl in &program.declarations {
                                if let Declaration::Sub(sub_decl) = decl {
                                    if let Some(Statement::Dim(dim_decl)) = sub_decl.body.first() {
                                        if let Some(expr) = &dim_decl.initializer {
                                            parts.push(expr.clone());
                                            found = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if !found {
                                parts.push(Expression::Variable(Identifier::new(expr_text.trim())));
                            }
                        }
                        Err(_) => {
                            // Fallback: treat as simple variable reference
                            parts.push(Expression::Variable(Identifier::new(expr_text.trim())));
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
                parts.push(Expression::StringLiteral(current_text));
            }
            
            // Build concatenation chain
            if parts.is_empty() {
                Ok(Expression::StringLiteral(String::new()))
            } else if parts.len() == 1 {
                Ok(parts.into_iter().next().unwrap())
            } else {
                let mut result = parts.remove(0);
                for part in parts {
                    result = Expression::Concatenate(
                        Box::new(result),
                        Box::new(part),
                    );
                }
                Ok(result)
            }
        }
        Rule::string_literal => {
            let s = pair.as_str();
            // Strip outer quotes, then unescape VB-style doubled quotes ("" -> ")
            let inner = s[1..s.len()-1].replace("\"\"", "\"");
            Ok(Expression::StringLiteral(inner))
        }
        Rule::numeric_literal => {
            let s = pair.as_str();
            // Strip type suffixes: F, D, L, R, S, US, UI, UL, !, #, @, %
            let s = s.trim_end_matches(|c: char| c.is_ascii_alphabetic() || c == '!' || c == '#' || c == '@' || c == '%');
            if s.contains('.') {
                Ok(Expression::DoubleLiteral(s.parse().unwrap_or(0.0)))
            } else {
                Ok(Expression::IntegerLiteral(s.parse().unwrap_or(0)))
            }
        }
        Rule::boolean_literal => {
            Ok(Expression::BooleanLiteral(pair.as_str().to_lowercase() == "true"))
        }
        Rule::array_literal => {
            parse_array_literal(pair)
        }
        Rule::date_literal => {
            let s = pair.as_str();
            // Strip the surrounding # delimiters
            let inner = s[1..s.len()-1].trim().to_string();
            Ok(Expression::DateLiteral(inner))
        }
        Rule::nothing_literal => Ok(Expression::Nothing),
        Rule::new_expression => {
            let mut inner = pair.into_inner();
            let id_pair = inner.next().unwrap();
            let mut class_name = id_pair.as_str().to_string();
            // Check for generic_suffix: List(Of String) -> "List(Of String)"
            let mut args = Vec::new();
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
                        return Ok(Expression::NewFromInitializer(
                            Identifier::new(&class_name),
                            args,
                            elements,
                        ));
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
                        return Ok(Expression::NewWithInitializer(
                            Identifier::new(&class_name),
                            args,
                            members,
                        ));
                    }
                    _ => {}
                }
            }
            // If there's an array initializer, return an ArrayLiteral instead of New
            if let Some(elements) = array_init {
                Ok(Expression::ArrayLiteral(elements))
            } else {
                Ok(Expression::New(Identifier::new(&class_name), args))
            }
        }
        Rule::if_expression => {
            let mut inner = pair.into_inner();
            let first = parse_expression(inner.next().unwrap())?;
            let second = parse_expression(inner.next().unwrap())?;
            let third = inner.next().map(|p| parse_expression(p)).transpose()?;
            Ok(Expression::IfExpression(
                Box::new(first),
                Box::new(second),
                third.map(Box::new),
            ))
        }
        Rule::addressof_expr => {
            let inner = pair.into_inner();
            let mut name = String::new();
            for p in inner {
                if p.as_rule() == Rule::dotted_identifier {
                    name = p.as_str().to_string();
                }
            }
            Ok(Expression::AddressOf(name))
        }
        Rule::me_keyword => {
            Ok(Expression::Me)
        }
        Rule::dot_call_statement => {
            // .Method(args) or .obj.Method(args) inside With block
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
                return Err(ParseError::Custom("dot_call needs at least one identifier".to_string()));
            }
            let method_name = Identifier::new(identifiers.last().unwrap().clone());
            let mut expr = Expression::WithTarget;
            for i in 0..identifiers.len() - 1 {
                expr = Expression::MemberAccess(Box::new(expr), Identifier::new(identifiers[i].clone()));
            }
            Ok(Expression::MethodCall(Box::new(expr), method_name, arguments))
        }
        Rule::dot_member_access => {
            // .prop or .obj.prop inside With block
            let inner = pair.into_inner();
            let mut expr = Expression::WithTarget;
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::MemberAccess(Box::new(expr), Identifier::new(p.as_str()));
                }
            }
            Ok(expr)
        }
        Rule::me_member_access => {
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut expr = Expression::Me;
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::MemberAccess(Box::new(expr), Identifier::new(p.as_str()));
                }
            }
            Ok(expr)
        }
        Rule::mybase_member_access => {
            // MyBase.Property
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut expr = Expression::MyBase;
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::MemberAccess(Box::new(expr), Identifier::new(p.as_str()));
                }
            }
            Ok(expr)
        }
        Rule::me_member_call => {
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::me_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err(ParseError::Custom("me_member_call needs at least one identifier".to_string()));
            }

            // Last identifier is the method name
            let method_name = Identifier::new(identifiers.last().unwrap().clone());

            // Build object expression: Me.a.b... (all except last)
            let mut expr = Expression::Me;
            for i in 0..identifiers.len() - 1 {
                expr = Expression::MemberAccess(Box::new(expr), Identifier::new(identifiers[i].clone()));
            }

            Ok(Expression::MethodCall(Box::new(expr), method_name, arguments))
        }
        Rule::mybase_member_call => {
            // MyBase.Method()
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::mybase_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err(ParseError::Custom("mybase_member_call needs at least one identifier".to_string()));
            }

            let method_name = Identifier::new(identifiers.last().unwrap().clone());
            let mut expr = Expression::MyBase;
            for i in 0..identifiers.len() - 1 {
                expr = Expression::MemberAccess(Box::new(expr), Identifier::new(identifiers[i].clone()));
            }

            Ok(Expression::MethodCall(Box::new(expr), method_name, arguments))
        }
        _ => Err(ParseError::UnexpectedRule(pair.as_rule())),
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> ParseResult<Expression> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinaryOp::Add,
                    "-" => BinaryOp::Subtract,
                    "*" => BinaryOp::Multiply,
                    "/" => BinaryOp::Divide,
                    "\\" => BinaryOp::IntegerDivide,
                    "mod" => BinaryOp::Modulo,
                    "^" => BinaryOp::Exponent,
                    "&" => BinaryOp::Concatenate,
                    "=" => BinaryOp::Equal,
                    "<>" => BinaryOp::NotEqual,
                    "<" => BinaryOp::LessThan,
                    "<=" => BinaryOp::LessThanOrEqual,
                    ">" => BinaryOp::GreaterThan,
                    ">=" => BinaryOp::GreaterThanOrEqual,
                    "and" => BinaryOp::And,
                    "andalso" => BinaryOp::AndAlso,
                    "or" => BinaryOp::Or,
                    "orelse" => BinaryOp::OrElse,
                    "xor" => BinaryOp::Xor,
                    "<<" => BinaryOp::BitShiftLeft,
                    ">>" => BinaryOp::BitShiftRight,
                    "is" => BinaryOp::Is,
                    "isnot" => BinaryOp::IsNot,
                    "like" => BinaryOp::Like,
                    _ => return Err(ParseError::Custom(format!("Unknown operator: {}", op_pair.as_str()))),
                }
            }
            _ => return Ok(left), // Should not happen with current grammar
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::binary(op, left, right);
    }

    Ok(left)
}

fn parse_unary_expression(pair: Pair<Rule>) -> ParseResult<Expression> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();

    match first.as_rule() {
        Rule::not_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::Not(Box::new(operand)))
        }
        Rule::neg_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::Negate(Box::new(operand)))
        }
        Rule::await_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::Await(Box::new(operand)))
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

fn parse_postfix_expression(pair: Pair<Rule>) -> ParseResult<Expression> {
    let mut inner = pair.into_inner();
    let primary = inner.next().unwrap();
    let mut expr = parse_expression(primary)?;

    // Apply member_chain postfix operations
    for chain in inner {
        expr = parse_member_chain_node(chain, expr)?;
    }

    Ok(expr)
}

fn parse_member_chain_node(chain: Pair<Rule>, expr: Expression) -> ParseResult<Expression> {
    match chain.as_rule() {
        Rule::member_chain_call => {
            let mut chain_inner = chain.into_inner();
            let name = chain_inner.next().unwrap().as_str();
            let arguments = if let Some(arg_list) = chain_inner.next() {
                parse_argument_list(arg_list)?
            } else {
                vec![]
            };
            Ok(Expression::MethodCall(Box::new(expr), Identifier::new(name), arguments))
        }
        Rule::member_chain_access => {
            let name = chain.into_inner().next().unwrap().as_str();
            Ok(Expression::MemberAccess(Box::new(expr), Identifier::new(name)))
        }
        Rule::member_chain => {
            let inner_chain = chain.into_inner().next().unwrap();
            parse_member_chain_node(inner_chain, expr)
        }
        _ => Ok(expr),
    }
}

fn parse_argument_list(pair: Pair<Rule>) -> ParseResult<Vec<Expression>> {
    pair.into_inner().map(parse_expression).collect()
}

fn parse_try_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
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
                // finally_block -> try_body
                // "Finally" keyword is hidden/atomic?
                // Re-check grammar: finally_block = { ^"Finally" ~ (NEWLINE | EOI) ~ try_body }
                // Inner of finally_block should contain try_body.
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
    
    Ok(Statement::Try { body, catches, finally })
}

fn parse_catch_block(pair: Pair<Rule>) -> ParseResult<CatchBlock> {
    let inner = pair.into_inner();
    let mut variable = None;
    let mut when_clause = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                let name = Identifier::new(p.as_str());
                variable = Some((name, None)); 
            }
            Rule::type_name => {
                if let Some((name, _)) = variable {
                     variable = Some((name, Some(VBType::from_str(p.as_str()))));
                }
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
    
    Ok(CatchBlock { variable, when_clause, body })
}

fn parse_continue_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let text = pair.as_str().to_lowercase();
    let continue_type = if text.contains("do") {
        ContinueType::Do
    } else if text.contains("for") {
        ContinueType::For
    } else {
        ContinueType::While
    };
    
    Ok(Statement::Continue(continue_type))

}

fn parse_lambda_expression(pair: Pair<Rule>) -> ParseResult<Expression> {
    let text = pair.as_str().trim_start();
    let is_function = text.to_lowercase().starts_with("function");
    
    let mut inner = pair.into_inner();
    let mut params = Vec::new();

    let mut next_pair = inner.next().ok_or_else(|| ParseError::Custom("Lambda missing body".to_string()))?;
    
    if next_pair.as_rule() == Rule::param_list {
        params = parse_param_list(next_pair)?;
        next_pair = inner.next().ok_or_else(|| ParseError::Custom("Lambda missing body".to_string()))?;
    }
    
    if is_function {
        Ok(Expression::Lambda {
            params,
            body: Box::new(LambdaBody::Expression(Box::new(parse_expression(next_pair)?)))
        })
    } else {
        Ok(Expression::Lambda {
            params,
            body: Box::new(LambdaBody::Statement(Box::new(parse_statement(next_pair)?)))
        })
    }
}

fn parse_block_body(pair: Pair<Rule>) -> ParseResult<Vec<Statement>> {
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

fn parse_for_each_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let variable = Identifier::new(inner.next().unwrap().as_str());
    let mut collection = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_name => {} // Skip "As Type" — we don't store it in ForEach AST
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
            _ => {}
        }
    }

    Ok(Statement::ForEach {
        variable,
        collection: collection.ok_or_else(|| ParseError::Custom("For Each missing collection".to_string()))?,
        body,
    })
}

fn parse_with_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
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

    Ok(Statement::With { object, body })
}

fn parse_using_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let variable = Identifier::new(inner.next().unwrap().as_str());
    
    // Collect remaining pairs
    let remaining: Vec<_> = inner.collect();
    
    // Find the expression (resource)
    let mut resource_expr = None;
    let mut body_start_idx = 0;
    
    for (idx, p) in remaining.iter().enumerate() {
        match p.as_rule() {
            Rule::type_name => {}, // Skip type annotation
            Rule::expression => {
                resource_expr = Some(parse_expression(p.clone())?);
                body_start_idx = idx + 1;
                break;
            }
            _ => {}
        }
    }
    
    let resource = resource_expr.ok_or_else(|| ParseError::Custom("Using statement missing resource expression".to_string()))?;
    
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

    Ok(Statement::Using { variable, resource, body })
}

fn parse_enum_decl(pair: Pair<Rule>) -> ParseResult<EnumDecl> {
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = Identifier::new("");
    let mut members = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                let text = p.as_str().to_lowercase();
                match text.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Friend,
                    _ => name = Identifier::new(p.as_str()),
                }
            }
            Rule::enum_member => {
                let mut member_inner = p.into_inner();
                let member_name = Identifier::new(member_inner.next().unwrap().as_str());
                let value = member_inner
                    .find(|e| e.as_rule() == Rule::expression)
                    .map(|e| parse_expression(e))
                    .transpose()?;
                members.push(EnumMember { name: member_name, value });
            }
            Rule::enum_end | Rule::NEWLINE => {}
            _ => {}
        }
    }

    Ok(EnumDecl { visibility, name, members })
}

fn parse_single_line_if(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let condition = parse_expression(inner.next().unwrap())?;

    // Parse then body
    let then_body_pair = inner.next().unwrap(); // single_line_then_body
    let mut then_branch = Vec::new();
    for stmt_pair in then_body_pair.into_inner() {
        then_branch.push(parse_statement(stmt_pair)?);
    }

    // Parse optional else body
    let else_branch = if let Some(else_body_pair) = inner.next() {
        let mut else_stmts = Vec::new();
        for stmt_pair in else_body_pair.into_inner() {
            else_stmts.push(parse_statement(stmt_pair)?);
        }
        Some(else_stmts)
    } else {
        None
    };

    Ok(Statement::If {
        condition,
        then_branch,
        elseif_branches: Vec::new(),
        else_branch,
    })
}

fn parse_field_decl(pair: Pair<Rule>) -> ParseResult<VariableDecl> {
    let mut field_name = Identifier::new("");
    let mut field_type = None;
    let mut field_init = None;
    let mut field_bounds = None;
    let mut is_new = false;
    let mut ctor_args: Vec<Expression> = Vec::new();
    
    for fp in pair.into_inner() {
        match fp.as_rule() {
            Rule::withevents_keyword => {} 
            Rule::visibility_modifier | Rule::partial_keyword => {} // modifiers handled by caller
            Rule::dim_new_keyword => { is_new = true; }
            Rule::identifier => field_name = Identifier::new(fp.as_str()),
            Rule::type_name => field_type = Some(VBType::from_str(fp.as_str())),
            Rule::array_bounds => {
                let bounds: Vec<Expression> = fp.into_inner()
                    .map(|e| parse_expression(e))
                    .collect::<Result<_, _>>()?;
                field_bounds = Some(bounds);
            }
            Rule::argument_list => {
                for arg_pair in fp.into_inner() {
                    if arg_pair.as_rule() == Rule::expression {
                        ctor_args.push(parse_expression(arg_pair)?);
                    }
                }
            }
            Rule::expression => field_init = Some(parse_expression(fp)?),
            Rule::array_literal => field_init = Some(parse_expression(fp)?),
            _ => {}
        }
    }
    
    // Handle "As New Type" syntax
    if is_new && field_init.is_none() {
        if let Some(t) = &field_type {
            field_init = Some(Expression::New(Identifier::new(t.to_string()), ctor_args));
        }
    }
    
    Ok(VariableDecl {
        name: field_name,
        var_type: field_type,
        array_bounds: field_bounds,
        initializer: field_init,
    })
}

fn parse_open_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let file_path = parse_expression(inner.next().unwrap())?;
    let mode_pair = inner.next().unwrap(); // file_mode
    let mode = match mode_pair.as_str().to_lowercase().as_str() {
        "input" => FileOpenMode::Input,
        "output" => FileOpenMode::Output,
        "append" => FileOpenMode::Append,
        "binary" => FileOpenMode::Binary,
        "random" => FileOpenMode::Random,
        _ => return Err(ParseError::Custom(format!("Unknown file mode: {}", mode_pair.as_str()))),
    };
    let file_number = parse_expression(inner.next().unwrap())?;
    Ok(Statement::Open { file_path, mode, file_number })
}

fn parse_close_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    // Close with no arguments closes all files
    let file_number = inner.next().map(|p| parse_expression(p)).transpose()?;
    Ok(Statement::CloseFile { file_number })
}

fn parse_print_file_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner.next()
        .map(parse_argument_list)
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::PrintFile { file_number, items, newline: true })
}

fn parse_write_file_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner.next()
        .map(parse_argument_list)
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::WriteFile { file_number, items })
}

fn parse_input_file_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let mut variables = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::identifier {
            variables.push(Identifier::new(p.as_str()));
        }
    }
    Ok(Statement::InputFile { file_number, variables })
}

fn parse_line_input_statement(pair: Pair<Rule>) -> ParseResult<Statement> {
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let variable = Identifier::new(inner.next().unwrap().as_str());
    Ok(Statement::LineInput { file_number, variable })
}
fn parse_select_statement(pair: Pair<Rule>) -> Result<Statement, ParseError> {
    let mut inner = pair.into_inner();
    let test_expr = parse_expression(inner.next().unwrap())?;
    
    let mut cases = Vec::new();
    let mut else_block = None;

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
                            // Can be simple value or Range (Expr To Expr)
                            let expr1 = parse_expression(first)?;
                            if let Some(next) = cond_inner.next() {
                                // removing "To" keyword if present in grammar structure, check rule
                                // Grammar: expression ~ ^"To" ~ expression
                                // The layout of case_condition rule:
                                // case_condition = {
                                //     ^"Is" ~ comp_op ~ expression           // Case Is > 10
                                //     | expression ~ ^"To" ~ expression      // Case 1 To 10
                                //     | expression                           // Case 5
                                // }
                                // If there's a second expression, it's a range
                                let expr2 = parse_expression(next)?;
                                CaseCondition::Range { from: expr1, to: expr2 }
                            } else {
                                CaseCondition::Value(expr1)
                            }
                        }
                        Rule::comp_op => {
                            // Is <op> Expression
                            let op = match first.as_str() {
                                "=" => CompOp::Equal,
                                "<>" => CompOp::NotEqual,
                                "<" => CompOp::LessThan,
                                "<=" => CompOp::LessThanOrEqual,
                                ">" => CompOp::GreaterThan,
                                ">=" => CompOp::GreaterThanOrEqual,
                                _ => return Err(ParseError::Custom(format!("Unknown comparison operator: {}", first.as_str()))),
                            };
                            let expr = parse_expression(cond_inner.next().unwrap())?;
                            CaseCondition::Comparison { op, expr }
                        }
                        _ => return Err(ParseError::Custom(format!("Unexpected rule in case condition: {:?}", first.as_rule()))),
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
                cases.push(CaseBlock { conditions, body });
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
                else_block = Some(body);
            }
            Rule::select_end => {}
            _ => {}
        }
    }

    Ok(Statement::Select {
        test_expr,
        cases,
        else_block,
    })
}
