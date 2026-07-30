use pest::iterators::Pair;
use pest::Parser;
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_ast::*;

use super::{KotlinParser, Rule};

static NEXT_TMP_ID: AtomicUsize = AtomicUsize::new(1);

fn gen_tmp_name() -> String {
    format!("__kt_tmp_{}", NEXT_TMP_ID.fetch_add(1, Ordering::Relaxed))
}

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs = KotlinParser::parse(Rule::program, source)
        .map_err(|e| format!("Kotlin parse error: {}", e))?;

    let root = pairs.next().ok_or_else(|| "Empty parse result".to_string())?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    for pair in root.into_inner() {
        match pair.as_rule() {
            Rule::package_decl => {}
            Rule::import_decl => {
                if let Some(imp) = walk_import(pair) {
                    imports.push(imp);
                }
            }
            Rule::top_level_decl => {
                let mut label_name = None;
                for inner in pair.into_inner() {
                    match inner.as_rule() {
                        Rule::label_decl => label_name = Some(inner.as_str().trim_end_matches('@').to_string()),
                        Rule::typealias_decl => {
                            if let Some(stmt) = walk_typealias(inner) {
                                body.push(stmt);
                            }
                        }
                        _ => {
                            if let Some(stmt) = walk_statement(inner) {
                                if let Some(lbl) = label_name.take() {
                                    body.push(Statement::new(StmtKind::Labeled {
                                        label: lbl,
                                        body: Box::new(stmt),
                                    }));
                                } else {
                                    body.push(stmt);
                                }
                            }
                        }
                    }
                }
            }
            Rule::EOI => {}
            _ => {}
        }
    }

    Ok(Module {
        name: "main".to_string(),
        language: Lang::Kotlin,
        body,
        imports,
    })
}

fn walk_typealias(_pair: Pair<Rule>) -> Option<Statement> {
    Some(Statement::new(StmtKind::Empty))
}

fn walk_annotation(pair: Pair<Rule>) -> Expression {
    let mut type_name = String::new();
    let mut args = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_ref => type_name = inner.as_str().to_string(),
            Rule::arg_list => {
                for arg_p in inner.into_inner() {
                    let mut arg_expr = None;
                    let mut arg_name = None;
                    for sub in arg_p.into_inner() {
                        match sub.as_rule() {
                            Rule::identifier => arg_name = Some(sub.as_str().to_string()),
                            Rule::expr => arg_expr = Some(walk_expr(sub)),
                            _ => {}
                        }
                    }
                    if let Some(ae) = arg_expr {
                        args.push(Argument {
                            value: ae,
                            name: arg_name,
                            by_ref: false,
                            spread: false,
                        });
                    }
                }
            }
            _ => {}
        }
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&type_name)),
        args,
        optional: false,
    })
}

fn walk_import(pair: Pair<Rule>) -> Option<Import> {
    let mut path = String::new();
    let mut alias = None;
    let mut is_wildcard = false;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::dotted_name => path = inner.as_str().to_string(),
            Rule::identifier => alias = Some(inner.as_str().to_string()),
            _ => {
                if inner.as_str() == ".*" || inner.as_str() == "*" {
                    is_wildcard = true;
                }
            }
        }
    }

    let kind = if is_wildcard {
        ImportKind::Wildcard { path, alias }
    } else {
        ImportKind::Simple { path, alias }
    };

    Some(Import {
        kind,
        span: Span::default(),
    })
}

fn walk_statement(pair: Pair<Rule>) -> Option<Statement> {
    let mut label_name = None;

    let inner_pair = if pair.as_rule() == Rule::statement {
        let mut inner_iter = pair.into_inner();
        let first = inner_iter.next()?;
        if first.as_rule() == Rule::label_decl {
            label_name = Some(first.as_str().trim_end_matches('@').to_string());
            inner_iter.next()?
        } else {
            first
        }
    } else {
        pair
    };

    let stmt = match inner_pair.as_rule() {
        Rule::typealias_decl => walk_typealias(inner_pair),
        Rule::interface_decl => walk_interface_decl(inner_pair),
        Rule::enum_decl => walk_enum_decl(inner_pair),
        Rule::destructuring_decl => walk_destructuring_decl(inner_pair),
        Rule::function_decl => walk_function_decl(inner_pair),
        Rule::var_decl => walk_var_decl(inner_pair),
        Rule::class_decl => walk_class_decl(inner_pair),
        Rule::object_decl => walk_object_decl(inner_pair),
        Rule::if_expr => walk_if_stmt(inner_pair),
        Rule::when_expr => walk_when_stmt(inner_pair),
        Rule::try_expr => walk_try_stmt(inner_pair),
        Rule::throw_stmt => {
            let expr = inner_pair.into_inner().find(|p| p.as_rule() == Rule::expr).map(walk_expr);
            Some(Statement::new(StmtKind::Throw { expr, cause: None }))
        }
        Rule::for_stmt => walk_for_stmt(inner_pair),
        Rule::while_stmt => walk_while_stmt(inner_pair),
        Rule::do_while_stmt => walk_do_while_stmt(inner_pair),
        Rule::return_stmt => {
            let mut ret_expr = None;
            for rsub in inner_pair.into_inner() {
                if rsub.as_rule() == Rule::expr {
                    ret_expr = Some(walk_expr(rsub));
                }
            }
            Some(Statement::new(StmtKind::Return(ret_expr)))
        }
        Rule::break_stmt => {
            let mut lbl = None;
            for bsub in inner_pair.into_inner() {
                if bsub.as_rule() == Rule::identifier {
                    lbl = Some(bsub.as_str().to_string());
                }
            }
            let target = lbl.map(BreakTarget::Label).unwrap_or(BreakTarget::Implicit);
            Some(Statement::new(StmtKind::Break(target)))
        }
        Rule::continue_stmt => {
            let mut lbl = None;
            for csub in inner_pair.into_inner() {
                if csub.as_rule() == Rule::identifier {
                    lbl = Some(csub.as_str().to_string());
                }
            }
            let target = lbl.map(ContinueTarget::Label).unwrap_or(ContinueTarget::Implicit);
            Some(Statement::new(StmtKind::Continue(target)))
        }
        Rule::expr_stmt => {
            let expr_pair = inner_pair.into_inner().next()?;
            let expr = walk_expr(expr_pair);
            Some(Statement::new(StmtKind::Expr(expr)))
        }
        Rule::expr => {
            let expr = walk_expr(inner_pair);
            Some(Statement::new(StmtKind::Expr(expr)))
        }
        _ => None,
    };

    match (stmt, label_name) {
        (Some(s), Some(lbl)) => Some(Statement::new(StmtKind::Labeled {
            label: lbl,
            body: Box::new(s),
        })),
        (other, _) => other,
    }
}

fn walk_interface_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();
    let mut decorators = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::identifier => {
                if name.is_empty() {
                    name = inner.as_str().to_string();
                }
            }
            Rule::inheritance_list => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::inheritance_specifier {
                        for sub in spec.into_inner() {
                            if sub.as_rule() == Rule::type_ref {
                                parents.push(sub.as_str().to_string());
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            match inner_member.as_rule() {
                                Rule::function_decl => {
                                    if let Some(stmt) = walk_function_decl(inner_member) {
                                        if let StmtKind::FunctionDecl { name, params, return_type, is_sub, .. } = stmt.kind {
                                            members.push(InterfaceMember::Method {
                                                name,
                                                params,
                                                return_type,
                                                is_sub,
                                                signature_source: None,
                                            });
                                        }
                                    }
                                }
                                Rule::var_decl => {
                                    if let Some(stmt) = walk_var_decl(inner_member) {
                                        if let StmtKind::VarDecl { declarations, kind } = stmt.kind {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(pname) = decl.pattern {
                                                    members.push(InterfaceMember::Property {
                                                        name: pname,
                                                        type_hint: decl.type_hint,
                                                        is_readonly: kind == VarDeclKind::Const,
                                                        is_writeonly: false,
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::InterfaceDecl {
        name,
        parents,
        members,
        decorators,
    }))
}

fn walk_enum_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut members = Vec::new();
    let mut body_members = Vec::new();
    let mut decorators = Vec::new();
    let mut entry_idx = 0i64;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::identifier => {
                if name.is_empty() {
                    name = inner.as_str().to_string();
                }
            }
            Rule::enum_entry => {
                let mut em_name = String::new();
                let mut ctor_args = Vec::new();
                for esub in inner.into_inner() {
                    match esub.as_rule() {
                        Rule::identifier => em_name = esub.as_str().to_string(),
                        Rule::arg_list => {
                            for arg_p in esub.into_inner() {
                                for e in arg_p.into_inner() {
                                    if e.as_rule() == Rule::expr {
                                        ctor_args.push(walk_expr(e));
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if !em_name.is_empty() {
                    let val_expr = if let Some(first_arg) = ctor_args.first() {
                        Some(first_arg.clone())
                    } else {
                        Some(Expression::int(entry_idx))
                    };
                    entry_idx += 1;
                    members.push(EnumMember {
                        name: em_name,
                        value: val_expr,
                        constructor_args: ctor_args,
                    });
                }
            }
            Rule::class_member => {
                if let Some(inner_member) = inner.into_inner().next() {
                    if inner_member.as_rule() == Rule::function_decl {
                        if let Some(stmt) = walk_function_decl(inner_member) {
                            body_members.push(ClassMember::Method(Box::new(stmt)));
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::EnumDecl {
        name,
        members,
        visibility: Visibility::Public,
        is_flags: false,
        backing_type: None,
        interfaces: vec![],
        body_members,
        decorators,
    }))
}

fn walk_destructuring_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut is_readonly = false;
    let mut names = Vec::new();
    let mut init = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::val_kw => is_readonly = true,
            Rule::var_kw => is_readonly = false,
            Rule::identifier => names.push(inner.as_str().to_string()),
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    if let Some(init_expr) = init {
        let tmp_name = gen_tmp_name();
        let decl_kind = if is_readonly { VarDeclKind::Const } else { VarDeclKind::Var };

        let mut stmts = vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(tmp_name.clone()),
                    type_hint: None,
                    init: Some(init_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: decl_kind.clone(),
            })
        ];

        for (idx, name) in names.into_iter().enumerate() {
            let read_expr = Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&tmp_name)),
                index: Box::new(Expression::int(idx as i64)),
                null_safe: false,
            });
            stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: decl_kind.clone(),
            }));
        }

        Some(Statement::new(StmtKind::Block(stmts)))
    } else {
        let elems = names.into_iter().map(|n| ArrayPatternElem::Pattern(BindingPattern::Ident(n), None)).collect();
        Some(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Array(elems),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: if is_readonly { VarDeclKind::Const } else { VarDeclKind::Var },
        }))
    }
}

fn walk_try_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block => {
                if body.is_empty() {
                    body = walk_block_statements(inner);
                }
            }
            Rule::catch_clause => {
                let mut param_name = "e".to_string();
                let mut type_hint = None;
                let mut catch_block_stmts = Vec::new();
                for csub in inner.into_inner() {
                    match csub.as_rule() {
                        Rule::identifier => param_name = csub.as_str().to_string(),
                        Rule::type_ref => type_hint = Some(csub.as_str().to_string()),
                        Rule::block => catch_block_stmts = walk_block_statements(csub),
                        _ => {}
                    }
                }
                let types = match type_hint.as_deref() {
                    Some("Exception") | Some("Throwable") | None => vec![],
                    Some(t) => vec![t.to_string()],
                };
                catches.push(CatchClause {
                    types,
                    var_name: Some(param_name),
                    stack_var: None,
                    body: catch_block_stmts,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                for fsub in inner.into_inner() {
                    if fsub.as_rule() == Rule::block {
                        finally = Some(walk_block_statements(fsub));
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    }))
}

fn walk_function_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut receiver_type: Option<String> = None;
    let mut params = Vec::new();
    let mut return_type = None;
    let mut body = Vec::new();

    let mut is_abstract = false;
    let mut visibility = Visibility::Public;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::modifier => {
                let m_str = inner.as_str();
                if m_str == "abstract" {
                    is_abstract = true;
                } else if m_str == "private" {
                    visibility = Visibility::Private;
                } else if m_str == "protected" {
                    visibility = Visibility::Protected;
                }
            }
            Rule::receiver_prefix => {
                if let Some(id_p) = inner.into_inner().next() {
                    receiver_type = Some(id_p.as_str().to_string());
                }
            }
            Rule::type_ref => {
                if return_type.is_none() && !name.is_empty() {
                    return_type = Some(inner.as_str().to_string());
                }
            }
            Rule::identifier => {
                name = inner.as_str().to_string();
            }
            Rule::parameter_list => {
                params = walk_parameter_list(inner);
            }
            Rule::function_body_expr => {
                if let Some(expr_pair) = inner.into_inner().next() {
                    let expr = walk_expr(expr_pair);
                    body.push(Statement::new(StmtKind::Return(Some(expr))));
                }
            }
            Rule::block => {
                body = walk_block_statements(inner);
            }
            _ => {}
        }
    }

    if receiver_type.is_some() {
        let mut ext_params = vec![Param {
            name: "this".to_string(),
            type_hint: receiver_type,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }];
        ext_params.extend(params);
        params = ext_params;
    }

    Some(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers {
            visibility,
            is_abstract,
            ..Default::default()
        },
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Vec<Param> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter {
            let mut is_rest = false;
            let mut name = String::new();
            let mut type_hint = None;
            let mut default = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::vararg_kw => is_rest = true,
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::type_ref => type_hint = Some(p.as_str().to_string()),
                    Rule::expr => default = Some(walk_expr(p)),
                    _ => {}
                }
            }
            let is_optional = default.is_some();
            let is_nullable = type_hint.as_ref().map_or(false, |t| t.ends_with('?'));
            params.push(Param {
                name,
                type_hint,
                default,
                pass_by: PassBy::Value,
                is_rest,
                is_kwargs: false,
                is_optional,
                is_nullable,
            });
        }
    }
    params
}

fn walk_var_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut is_readonly = false;
    let mut is_const = false;
    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::modifier => {
                if inner.as_str() == "const" {
                    is_const = true;
                }
            }
            Rule::val_kw => is_readonly = true,
            Rule::var_kw => is_readonly = false,
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::type_ref => type_hint = Some(inner.as_str().to_string()),
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds: None,
            with_events: false,
        }],
        kind: if is_const || is_readonly { VarDeclKind::Const } else { VarDeclKind::Var },
    }))
}

fn walk_class_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut is_interface = false;
    let mut is_abstract = false;
    let mut is_sealed = false;

    let mut parents = Vec::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();
    let mut base_args = None;
    let mut decorators = Vec::new();
    let mut init_stmts = Vec::new();

    let inner_pairs: Vec<_> = pair.into_inner().collect();

    for inner in &inner_pairs {
        if inner.as_rule() == Rule::inheritance_list {
            for spec in inner.clone().into_inner() {
                if spec.as_rule() == Rule::inheritance_specifier {
                    let mut parent_name = String::new();
                    let mut spec_base_args = Vec::new();
                    let mut by_expr = None;
                    for sub in spec.into_inner() {
                        match sub.as_rule() {
                            Rule::type_ref => parent_name = sub.as_str().to_string(),
                            Rule::arg_list => {
                                for arg_p in sub.into_inner() {
                                    let mut arg_expr = None;
                                    for e in arg_p.into_inner() {
                                        if e.as_rule() == Rule::expr {
                                            arg_expr = Some(walk_expr(e));
                                        }
                                    }
                                    if let Some(ae) = arg_expr {
                                        spec_base_args.push(ae);
                                    }
                                }
                            }
                            Rule::expr => by_expr = Some(sub.as_str().to_string()),
                            _ => {}
                        }
                    }
                    if !parent_name.is_empty() {
                        if by_expr.is_some() {
                            members.push(ClassMember::Augment(AugmentDecl {
                                from: parent_name.clone(),
                                via_field: by_expr,
                                adjustments: vec![],
                            }));
                        } else if parents.is_empty() && !is_interface {
                            parents.push(parent_name);
                            if !spec_base_args.is_empty() {
                                base_args = Some(spec_base_args);
                            }
                        } else {
                            interfaces.push(parent_name);
                        }
                    }
                }
            }
        }
    }

    for inner in inner_pairs {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::interface_kw => is_interface = true,
            Rule::modifier => {
                match inner.as_str() {
                    "abstract" => is_abstract = true,
                    "sealed" => is_sealed = true,
                    _ => {}
                }
            }
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::primary_constructor => {
                let mut ctor_params = Vec::new();
                let mut ctor_body = Vec::new();

                for param in inner.into_inner() {
                    if param.as_rule() == Rule::class_parameter {
                        let mut param_is_prop = false;
                        let mut is_readonly = false;
                        let mut pname = String::new();
                        let mut type_hint = None;
                        for p in param.into_inner() {
                            match p.as_rule() {
                                Rule::val_kw => {
                                    param_is_prop = true;
                                    is_readonly = true;
                                }
                                Rule::var_kw => {
                                    param_is_prop = true;
                                    is_readonly = false;
                                }
                                Rule::identifier => pname = p.as_str().to_string(),
                                Rule::type_ref => type_hint = Some(p.as_str().to_string()),
                                _ => {}
                            }
                        }
                        if !pname.is_empty() {
                            ctor_params.push(Param {
                                name: pname.clone(),
                                type_hint: type_hint.clone(),
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                            if param_is_prop {
                                members.push(ClassMember::Field {
                                    name: pname.clone(),
                                    type_hint: type_hint.clone(),
                                    init: None,
                                    modifiers: Modifiers {
                                        visibility: Visibility::Public,
                                        is_readonly,
                                        ..Default::default()
                                    },
                                    with_events: false,
                                    array_bounds: None,
                                });
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            field: pname.clone(),
                                            null_safe: false,
                                        })),
                                        value: Box::new(Expression::ident(&pname)),
                                    },
                                ))));
                            }
                        }
                    }
                }

                members.push(ClassMember::Constructor {
                    name: None,
                    params: ctor_params,
                    body: ctor_body,
                    base_args: base_args.clone(),
                    initializer_target: ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public,
                });
            }
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            match inner_member.as_rule() {
                                Rule::init_block => {
                                    if let Some(block_pair) = inner_member.into_inner().next() {
                                        init_stmts.extend(walk_block_statements(block_pair));
                                    }
                                }
                                Rule::secondary_constructor => {
                                    let mut s_params = Vec::new();
                                    let mut s_body = Vec::new();
                                    let mut s_target = ConstructorInitializerTarget::Base;
                                    let mut s_base_args = None;
                                    for sc in inner_member.into_inner() {
                                        match sc.as_rule() {
                                            Rule::parameter_list => s_params = walk_parameter_list(sc),
                                            Rule::this_kw => s_target = ConstructorInitializerTarget::This,
                                            Rule::super_kw => s_target = ConstructorInitializerTarget::Base,
                                            Rule::arg_list => {
                                                let mut bargs = Vec::new();
                                                for arg_p in sc.into_inner() {
                                                    for e in arg_p.into_inner() {
                                                        if e.as_rule() == Rule::expr {
                                                            bargs.push(walk_expr(e));
                                                        }
                                                    }
                                                }
                                                s_base_args = Some(bargs);
                                            }
                                            Rule::block => s_body = walk_block_statements(sc),
                                            _ => {}
                                        }
                                    }
                                    members.push(ClassMember::Constructor {
                                        name: None,
                                        params: s_params,
                                        body: s_body,
                                        base_args: s_base_args,
                                        initializer_target: s_target,
                                        visibility: Visibility::Public,
                                    });
                                }
                                Rule::class_decl | Rule::object_decl | Rule::interface_decl => {
                                    if let Some(stmt) = walk_statement(inner_member) {
                                        members.push(ClassMember::NestedType(Box::new(stmt)));
                                    }
                                }
                                Rule::function_decl => {
                                    if let Some(stmt) = walk_function_decl(inner_member) {
                                        members.push(ClassMember::Method(Box::new(stmt)));
                                    }
                                }
                                Rule::var_decl => {
                                    if let Some(stmt) = walk_var_decl(inner_member) {
                                        if let StmtKind::VarDecl { declarations, kind } = stmt.kind {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(fname) = decl.pattern {
                                                    if kind == VarDeclKind::Const {
                                                        if let Some(val_expr) = decl.init {
                                                            members.push(ClassMember::Const {
                                                                name: fname,
                                                                type_hint: decl.type_hint,
                                                                value: val_expr,
                                                                visibility: Visibility::Public,
                                                            });
                                                        }
                                                    } else {
                                                        members.push(ClassMember::Field {
                                                            name: fname,
                                                            type_hint: decl.type_hint,
                                                            init: decl.init,
                                                            modifiers: Modifiers {
                                                                visibility: Visibility::Public,
                                                                ..Default::default()
                                                            },
                                                            with_events: false,
                                                            array_bounds: None,
                                                        });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Rule::companion_object => {
                                    if let Some(stmt) = walk_object_decl(inner_member) {
                                        if let StmtKind::ClassDecl { members: comp_members, .. } = stmt.kind {
                                            for mut cm in comp_members {
                                                if let ClassMember::Method(ref mut mstmt) = cm {
                                                    if let StmtKind::FunctionDecl { ref mut modifiers, .. } = mstmt.kind {
                                                        modifiers.is_static = true;
                                                    }
                                                }
                                                members.push(cm);
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    if !init_stmts.is_empty() {
        for member in &mut members {
            if let ClassMember::Constructor { body, .. } = member {
                let mut combined = init_stmts.clone();
                combined.extend(std::mem::take(body));
                *body = combined;
            }
        }
    }

    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers {
            is_abstract: is_abstract || is_interface,
            is_sealed,
            ..Default::default()
        },
        decorators,
    }))
}

fn walk_object_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = "Companion".to_string();
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            if inner_member.as_rule() == Rule::function_decl {
                                if let Some(mut stmt) = walk_function_decl(inner_member) {
                                    if let StmtKind::FunctionDecl { ref mut modifiers, .. } = stmt.kind {
                                        modifiers.is_static = true;
                                    }
                                    members.push(ClassMember::Method(Box::new(stmt)));
                                }
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces: vec![],
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    }))
}

fn walk_block_statements(pair: Pair<Rule>) -> Vec<Statement> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if let Some(stmt) = walk_statement(inner) {
            stmts.push(stmt);
        }
    }
    stmts
}

fn walk_if_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut then_body = Vec::new();
    let mut else_body = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => {
                if then_body.is_empty() {
                    then_body = walk_block_statements(p);
                } else {
                    else_body = Some(walk_block_statements(p));
                }
            }
            Rule::statement => {
                if then_body.is_empty() {
                    if let Some(s) = walk_statement(p) {
                        then_body = vec![s];
                    }
                } else {
                    if let Some(s) = walk_statement(p) {
                        else_body = Some(vec![s]);
                    }
                }
            }
            Rule::if_expr => {
                if let Some(s) = walk_if_stmt(p) {
                    else_body = Some(vec![s]);
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::If {
        cond,
        then_body,
        elifs: vec![],
        else_body,
    }))
}

fn walk_when_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut disc = None;
    let mut entries = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => disc = Some(walk_expr(p)),
            Rule::when_entry => entries.push(p),
            _ => {}
        }
    }

    let mut cases = Vec::new();
    let mut default = None;

    for entry in entries {
        let mut entry_inner = entry.into_inner();
        let mut is_else = false;
        let mut cond_cases = Vec::new();
        let mut body_stmts = Vec::new();

        while let Some(p) = entry_inner.next() {
            match p.as_rule() {
                Rule::else_kw => is_else = true,
                Rule::when_condition => {
                    for csub in p.into_inner() {
                        match csub.as_rule() {
                            Rule::range_condition => {
                                let mut r_exprs = csub.into_inner();
                                if let (Some(e1), Some(e2)) = (r_exprs.next(), r_exprs.next()) {
                                    cond_cases.push(CaseCondition::Range {
                                        from: walk_expr(e1),
                                        to: walk_expr(e2),
                                    });
                                }
                            }
                            Rule::comparison_condition => {
                                let op_str = csub.as_str();
                                let mut c_inner = csub.into_inner();
                                let comp_op = if op_str.starts_with(">=") {
                                    ComparisonOp::GtEq
                                } else if op_str.starts_with("<=") {
                                    ComparisonOp::LtEq
                                } else if op_str.starts_with('>') {
                                    ComparisonOp::Gt
                                } else if op_str.starts_with('<') {
                                    ComparisonOp::Lt
                                } else if op_str.starts_with("!=") {
                                    ComparisonOp::NotEq
                                } else {
                                    ComparisonOp::Eq
                                };
                                if let Some(e) = c_inner.next() {
                                    cond_cases.push(CaseCondition::Comparison {
                                        op: comp_op,
                                        expr: walk_expr(e),
                                    });
                                }
                            }
                            Rule::expr => {
                                cond_cases.push(CaseCondition::Value(walk_expr(csub)));
                            }
                            _ => {}
                        }
                    }
                }
                Rule::block => body_stmts = walk_block_statements(p),
                Rule::statement => {
                    if let Some(s) = walk_statement(p) {
                        body_stmts.push(s);
                    }
                }
                _ => {}
            }
        }

        if is_else {
            default = Some(body_stmts);
        } else if !cond_cases.is_empty() {
            cases.push(SwitchCase {
                conditions: cond_cases,
                body: body_stmts,
            });
        }
    }

    let discriminator = disc.unwrap_or_else(|| Expression::bool(true));
    Some(Statement::new(StmtKind::Switch {
        expr: discriminator,
        cases,
        default,
    }))
}

fn walk_for_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut var_id = String::new();
    let mut destruct_names = Vec::new();
    let mut iter_expr = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => var_id = p.as_str().to_string(),
            Rule::for_destructure => {
                for dsub in p.into_inner() {
                    if dsub.as_rule() == Rule::identifier {
                        destruct_names.push(dsub.as_str().to_string());
                    }
                }
            }
            Rule::expr => iter_expr = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    if !destruct_names.is_empty() {
        let loop_tmp = gen_tmp_name();
        let mut prepended_stmts = Vec::new();
        for (idx, name) in destruct_names.into_iter().enumerate() {
            let read_expr = Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&loop_tmp)),
                index: Box::new(Expression::int(idx as i64)),
                null_safe: false,
            });
            prepended_stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }));
        }
        prepended_stmts.extend(body);
        body = prepended_stmts;
        var_id = loop_tmp;
    }

    Some(Statement::new(StmtKind::ForIn {
        var: var_id,
        key: None,
        iter: iter_expr,
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
}

fn walk_while_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::While {
        cond,
        body,
        else_body: None,
    }))
}

fn walk_do_while_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    }))
}

fn walk_lambda(pair: Pair<Rule>) -> Expression {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::lambda_params => {
                for lp in inner.into_inner() {
                    if lp.as_rule() == Rule::lambda_param {
                        let mut name = String::new();
                        let mut type_hint = None;
                        for lsub in lp.into_inner() {
                            match lsub.as_rule() {
                                Rule::identifier => name = lsub.as_str().to_string(),
                                Rule::type_ref => type_hint = Some(lsub.as_str().to_string()),
                                _ => {}
                            }
                        }
                        if !name.is_empty() {
                            let is_nullable = type_hint.as_ref().map_or(false, |t| t.ends_with('?'));
                            params.push(Param {
                                name,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable,
                            });
                        }
                    }
                }
            }
            Rule::statement => {
                if let Some(s) = walk_statement(inner) {
                    body.push(s);
                }
            }
            _ => {}
        }
    }

    if params.is_empty() {
        params.push(Param {
            name: "it".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }

    if let Some(last) = body.pop() {
        match last.kind {
            StmtKind::Expr(e) => {
                body.push(Statement::new(StmtKind::Return(Some(e))));
            }
            other => {
                body.push(Statement::new(other));
            }
        }
    }

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        captures: vec![],
        is_async: false,
    })
}

fn walk_expr(pair: Pair<Rule>) -> Expression {
    let rule = pair.as_rule();
    match rule {
        Rule::expr | Rule::assignment => {
            let mut inner = pair.into_inner();
            let first = inner.next().unwrap();
            if let Some(op_pair) = inner.next() {
                let val_pair = inner.next().unwrap();
                let lhs = walk_expr(first);
                let rhs = walk_expr(val_pair);
                let op_str = op_pair.as_str();
                if op_str == "=" {
                    Expression::new(ExprKind::Assign {
                        target: Box::new(lhs),
                        value: Box::new(rhs),
                    })
                } else {
                    let bin_op = match op_str {
                        "+=" => BinOp::Add,
                        "-=" => BinOp::Sub,
                        "*=" => BinOp::Mul,
                        "/=" => BinOp::Div,
                        "%=" => BinOp::Mod,
                        _ => BinOp::Add,
                    };
                    Expression::new(ExprKind::Assign {
                        target: Box::new(lhs.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: bin_op,
                            left: Box::new(lhs),
                            right: Box::new(rhs),
                        })),
                    })
                }
            } else {
                walk_expr(first)
            }
        }
        Rule::elvis => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(_op) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                current = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::logical_or => walk_binary_chain(pair, BinOp::Or),
        Rule::logical_and => walk_binary_chain(pair, BinOp::And),
        Rule::equality => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op = match op_pair.as_str() {
                    "==" | "===" => BinOp::Eq,
                    "!=" | "!==" => BinOp::NotEq,
                    _ => BinOp::Eq,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::comparison => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_pair = inner.next().unwrap();
                let op_str = op_pair.as_str();
                let type_str = next_pair.as_str().to_lowercase();
                current = match op_str {
                    "<" => Expression::new(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "<=" => Expression::new(ExprKind::Binary {
                        op: BinOp::LtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    ">" => Expression::new(ExprKind::Binary {
                        op: BinOp::Gt,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    ">=" => Expression::new(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "in" => Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "!in" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::In,
                            left: Box::new(current),
                            right: Box::new(walk_expr(next_pair)),
                        })),
                    }),
                    "is" => Expression::new(ExprKind::IsType {
                        expr: Box::new(current),
                        type_name: type_str,
                    }),
                    "!is" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::IsType {
                            expr: Box::new(current),
                            type_name: type_str,
                        })),
                    }),
                    _ => Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                };
            }
            current
        }
        Rule::range_expr => {
            let mut inner = pair.into_inner();
            let first = walk_expr(inner.next().unwrap());
            if let Some(_op) = inner.next() {
                let second = walk_expr(inner.next().unwrap());
                Expression::new(ExprKind::Range {
                    start: Box::new(first),
                    end: Box::new(second),
                    inclusive: true,
                })
            } else {
                first
            }
        }
        Rule::additive => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op = match op_pair.as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    _ => BinOp::Add,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::multiplicative => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op = match op_pair.as_str() {
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "%" => BinOp::Mod,
                    _ => BinOp::Mul,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::type_cast => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(_op_pair) = inner.next() {
                let target_type = inner.next().unwrap().as_str().to_string();
                current = Expression::new(ExprKind::Cast {
                    expr: Box::new(current),
                    type_name: target_type,
                });
            }
            current
        }
        Rule::infix_expr => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op_str = op_pair.as_str();
                if op_str == "to" {
                    current = Expression::new(ExprKind::Array(vec![
                        ArrayElement { value: current, spread: false },
                        ArrayElement { value: next_expr, spread: false },
                    ]));
                } else if op_str == "until" {
                    current = Expression::new(ExprKind::Range {
                        start: Box::new(current),
                        end: Box::new(next_expr),
                        inclusive: false,
                    });
                } else if op_str == "downTo" {
                    current = Expression::new(ExprKind::Range {
                        start: Box::new(current),
                        end: Box::new(next_expr),
                        inclusive: true,
                    });
                } else {
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(current),
                            field: op_str.to_string(),
                            null_safe: false,
                        })),
                        args: vec![Argument::positional(next_expr)],
                        optional: false,
                    });
                }
            }
            current
        }
        Rule::unary => {
            let mut inner = pair.into_inner();
            let mut ops = Vec::new();
            while let Some(p) = inner.next() {
                if p.as_rule() == Rule::prefix_op {
                    ops.push(p.as_str().to_string());
                } else {
                    let mut current = walk_expr(p);
                    for op in ops.into_iter().rev() {
                        let un_op = match op.as_str() {
                            "!" => UnaryOp::Not,
                            "-" => UnaryOp::Neg,
                            "+" => UnaryOp::Pos,
                            _ => UnaryOp::Not,
                        };
                        current = Expression::new(ExprKind::Unary {
                            op: un_op,
                            expr: Box::new(current),
                        });
                    }
                    return current;
                }
            }
            Expression::null()
        }
        Rule::postfix => {
            let mut inner = pair.into_inner();
            let primary_pair = inner.next().unwrap();
            let mut current = walk_expr(primary_pair);

            for suffix_pair in inner {
                let suffix_inner = suffix_pair.into_inner().next().unwrap();
                match suffix_inner.as_rule() {
                    Rule::call_suffix => {
                        let mut args = Vec::new();
                        for item in suffix_inner.into_inner() {
                            match item.as_rule() {
                                Rule::arg_list => {
                                    for arg_p in item.into_inner() {
                                        let mut arg_expr = None;
                                        let mut arg_name = None;
                                        let mut is_spread = false;
                                        for sub in arg_p.into_inner() {
                                            match sub.as_rule() {
                                                Rule::spread_op => is_spread = true,
                                                Rule::identifier => arg_name = Some(sub.as_str().to_string()),
                                                Rule::expr => arg_expr = Some(walk_expr(sub)),
                                                _ => {}
                                            }
                                        }
                                        if let Some(ae) = arg_expr {
                                            args.push(Argument {
                                                value: ae,
                                                name: arg_name,
                                                by_ref: false,
                                                spread: is_spread,
                                            });
                                        }
                                    }
                                }
                                Rule::lambda_literal => {
                                    args.push(Argument::positional(walk_lambda(item)));
                                }
                                _ => {}
                            }
                        }

                        let func_name = match &current.kind {
                            ExprKind::Ident(name) => Some(name.clone()),
                            _ => None,
                        };

                        if let Some(ref fn_name) = func_name {
                            if matches!(fn_name.as_str(), "Pair" | "Triple") {
                                let elems = args.into_iter().map(|a| ArrayElement { value: a.value, spread: false }).collect();
                                current = Expression::new(ExprKind::Array(elems));
                                continue;
                            } else if matches!(fn_name.as_str(), "mapOf" | "mutableMapOf" | "linkedMapOf" | "hashMapOf") {
                                let mut props = Vec::new();
                                for arg in args {
                                    if let ExprKind::Array(ref elems) = arg.value.kind {
                                        if elems.len() >= 2 {
                                            props.push(ObjectProperty {
                                                key: match &elems[0].value.kind {
                                                    ExprKind::String(s) => PropertyKey::Ident(s.clone()),
                                                    ExprKind::Ident(s) => PropertyKey::Ident(s.clone()),
                                                    _ => PropertyKey::Ident("key".to_string()),
                                                },
                                                value: elems[1].value.clone(),
                                                shorthand: false,
                                                computed: false,
                                            });
                                        }
                                    }
                                }
                                current = Expression::new(ExprKind::Object(props));
                                continue;
                            }
                        }

                        let is_class_name = match &current.kind {
                            ExprKind::Ident(name) => {
                                name.chars().next().map_or(false, |c| c.is_uppercase())
                                    && !matches!(name.as_str(), "Exception" | "IllegalArgumentException" | "IllegalStateException" | "NullPointerException" | "IndexOutOfBoundsException")
                            }
                            _ => false,
                        };

                        if is_class_name {
                            current = Expression::new(ExprKind::New {
                                class: Box::new(current),
                                args,
                            });
                        } else {
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(current),
                                args,
                                optional: false,
                            });
                        }
                    }
                    Rule::member_suffix => {
                        let field_id = suffix_inner.into_inner().next().unwrap().as_str().to_string();
                        if let ExprKind::Super = current.kind {
                            current = Expression::new(ExprKind::SuperCall {
                                method: Some(field_id),
                                args: vec![],
                            });
                        } else {
                            match field_id.as_str() {
                                "first" | "component1" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false,
                                    });
                                }
                                "second" | "component2" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(1)),
                                        null_safe: false,
                                    });
                                }
                                "third" | "component3" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(2)),
                                        null_safe: false,
                                    });
                                }
                                "size" => {
                                    current = Expression::new(ExprKind::Member {
                                        object: Box::new(current),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                }
                                _ => {
                                    current = Expression::new(ExprKind::Member {
                                        object: Box::new(current),
                                        field: field_id,
                                        null_safe: false,
                                    });
                                }
                            }
                        }
                    }
                    Rule::safe_call_suffix => {
                        let field_id = suffix_inner.into_inner().next().unwrap().as_str().to_string();
                        current = Expression::new(ExprKind::Member {
                            object: Box::new(current),
                            field: field_id,
                            null_safe: true,
                        });
                    }
                    Rule::index_suffix => {
                        let index_pair = suffix_inner.into_inner().next().unwrap();
                        let idx_expr = walk_expr(index_pair.into_inner().next().unwrap().into_inner().next().unwrap());
                        current = Expression::new(ExprKind::Index {
                            object: Box::new(current),
                            index: Box::new(idx_expr),
                            null_safe: false,
                        });
                    }
                    Rule::null_assert_suffix => {
                        current = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(current),
                        });
                    }
                    _ => {}
                }
            }
            current
        }
        Rule::primary => {
            let inner = pair.into_inner().next().unwrap();
            match inner.as_rule() {
                Rule::identifier => Expression::ident(inner.as_str()),
                Rule::literal => walk_literal(inner),
                Rule::this_kw => Expression::new(ExprKind::This),
                Rule::super_kw => Expression::new(ExprKind::Super),
                Rule::lambda_literal => walk_lambda(inner),
                Rule::object_expr => {
                    let mut parent = None;
                    let mut interfaces = Vec::new();
                    let mut members = Vec::new();
                    for osub in inner.into_inner() {
                        match osub.as_rule() {
                            Rule::inheritance_list => {
                                for spec in osub.into_inner() {
                                    if spec.as_rule() == Rule::inheritance_specifier {
                                        for sub in spec.into_inner() {
                                            if sub.as_rule() == Rule::type_ref {
                                                let tname = sub.as_str().to_string();
                                                if parent.is_none() {
                                                    parent = Some(Box::new(Expression::ident(&tname)));
                                                } else {
                                                    interfaces.push(tname);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            Rule::class_body => {
                                for member_pair in osub.into_inner() {
                                    if member_pair.as_rule() == Rule::class_member {
                                        if let Some(inner_member) = member_pair.into_inner().next() {
                                            if inner_member.as_rule() == Rule::function_decl {
                                                if let Some(stmt) = walk_function_decl(inner_member) {
                                                    members.push(ClassMember::Method(Box::new(stmt)));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    Expression::new(ExprKind::ClassExpr {
                        name: None,
                        parent,
                        interfaces,
                        members,
                    })
                }
                Rule::if_expr => {
                    let stmt = walk_if_stmt(inner).unwrap();
                    if let StmtKind::If { cond, then_body, else_body, .. } = stmt.kind {
                        let then_expr = then_body.into_iter().last().and_then(|s| match s.kind {
                            StmtKind::Expr(e) => Some(e),
                            StmtKind::Return(Some(e)) => Some(e),
                            _ => None,
                        }).unwrap_or_else(Expression::null);

                        let else_expr = else_body.unwrap_or_default().into_iter().last().and_then(|s| match s.kind {
                            StmtKind::Expr(e) => Some(e),
                            StmtKind::Return(Some(e)) => Some(e),
                            _ => None,
                        }).unwrap_or_else(Expression::null);

                        Expression::new(ExprKind::Ternary {
                            cond: Box::new(cond),
                            then: Box::new(then_expr),
                            else_: Box::new(else_expr),
                        })
                    } else {
                        Expression::null()
                    }
                }
                Rule::when_expr => {
                    let mut disc = None;
                    let mut entries = Vec::new();

                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::expr => disc = Some(walk_expr(p)),
                            Rule::when_entry => entries.push(p),
                            _ => {}
                        }
                    }

                    let mut arms = Vec::new();

                    for entry in entries {
                        let mut entry_inner = entry.into_inner();
                        let mut is_else = false;
                        let mut cond_exprs = Vec::new();
                        let mut body_expr = Expression::null();

                        while let Some(p) = entry_inner.next() {
                            match p.as_rule() {
                                Rule::else_kw => is_else = true,
                                Rule::when_condition => {
                                    for csub in p.into_inner() {
                                        if csub.as_rule() == Rule::expr {
                                            cond_exprs.push(walk_expr(csub));
                                        }
                                    }
                                }
                                Rule::block => {
                                    let stmts = walk_block_statements(p);
                                    if let Some(last) = stmts.into_iter().last() {
                                        body_expr = match last.kind {
                                            StmtKind::Expr(e) => e,
                                            StmtKind::Return(Some(e)) => e,
                                            _ => Expression::null(),
                                        };
                                    }
                                }
                                Rule::statement => {
                                    if let Some(s) = walk_statement(p) {
                                        body_expr = match s.kind {
                                            StmtKind::Expr(e) => e,
                                            StmtKind::Return(Some(e)) => e,
                                            _ => Expression::null(),
                                        };
                                    }
                                }
                                _ => {}
                            }
                        }

                        let conditions = if is_else {
                            None
                        } else {
                            Some(cond_exprs)
                        };

                        arms.push(MatchArm {
                            conditions,
                            body: body_expr,
                        });
                    }

                    Expression::new(ExprKind::Match {
                        subject: Box::new(disc.unwrap_or_else(|| Expression::bool(true))),
                        arms,
                    })
                }
                Rule::expr => walk_expr(inner),
                _ => Expression::null(),
            }
        }
        _ => Expression::null(),
    }
}

fn walk_binary_chain(pair: Pair<Rule>, op: BinOp) -> Expression {
    let mut inner = pair.into_inner();
    let mut current = walk_expr(inner.next().unwrap());
    while let Some(_op_pair) = inner.next() {
        let next_expr = walk_expr(inner.next().unwrap());
        current = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(current),
            right: Box::new(next_expr),
        });
    }
    current
}

fn walk_literal(pair: Pair<Rule>) -> Expression {
    let inner = pair.into_inner().next().unwrap();
    match inner.as_rule() {
        Rule::null_kw => Expression::null(),
        Rule::true_kw => Expression::bool(true),
        Rule::false_kw => Expression::bool(false),
        Rule::int_literal => {
            let s = inner.as_str().trim_end_matches(['L', 'l']);
            let val = s.parse::<i64>().unwrap_or(0);
            Expression::int(val)
        }
        Rule::float_literal => {
            let s = inner.as_str().trim_end_matches(['f', 'F']);
            let val = s.parse::<f64>().unwrap_or(0.0);
            Expression::float(val)
        }
        Rule::string_literal => walk_string_literal(inner),
        _ => Expression::null(),
    }
}

fn walk_string_literal(pair: Pair<Rule>) -> Expression {
    let mut parts = Vec::new();
    for content in pair.into_inner() {
        if content.as_rule() == Rule::string_content {
            if let Some(inner) = content.into_inner().next() {
                match inner.as_rule() {
                    Rule::str_text => {
                        parts.push(Expression::string(inner.as_str()));
                    }
                    Rule::str_escaped => {
                        let esc = match inner.as_str() {
                            "\\n" => "\n",
                            "\\t" => "\t",
                            "\\r" => "\r",
                            "\\\"" => "\"",
                            "\\\\" => "\\",
                            s => s,
                        };
                        parts.push(Expression::string(esc));
                    }
                    Rule::str_interpolated_var => {
                        if let Some(id_pair) = inner.into_inner().next() {
                            parts.push(Expression::ident(id_pair.as_str()));
                        }
                    }
                    Rule::str_interpolated_expr => {
                        if let Some(expr_pair) = inner.into_inner().next() {
                            parts.push(walk_expr(expr_pair));
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    if parts.is_empty() {
        Expression::string("")
    } else if parts.len() == 1 {
        parts.remove(0)
    } else {
        let mut iter = parts.into_iter();
        let mut acc = iter.next().unwrap();
        for p in iter {
            acc = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(acc),
                right: Box::new(p),
            });
        }
        acc
    }
}
