use crate::emitter::tostring::SET_MARKER;
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

/// True when `expr` is a bare dotted chain of identifiers (`java.util`), which
/// is what distinguishes a package-qualified type name from member access on a
/// value.
fn is_ident_chain(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(_) => true,
        ExprKind::Member { object, .. } => is_ident_chain(object),
        _ => false,
    }
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
        // Kotlin's import binds the SIMPLE NAME (`import java.time.Instant`
        // makes `Instant` mean `java.time.Instant`), so an import with no
        // `as` clause still carries an alias — its last segment. Without it
        // the import was inert: `Instant.parse(…)` resolved to nothing and
        // trapped "undefined is not callable". Java's `walk_import` has done
        // this all along; it is what populates `source_type_aliases`.
        let alias = alias.or_else(|| path.rsplit('.').next().map(str::to_string));
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
            Some(repeat_to_for_in(&expr).unwrap_or_else(|| Statement::new(StmtKind::Expr(expr))))
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

/// `repeat(n) { … }` -> the `for` loop it stands for.
///
/// Kotlin spells this control structure as a function, but it IS a loop: the
/// lambda runs `n` times and receives the 0-based index. Desugaring it here
/// rather than adapting it to a call is what puts `break` and `continue` inside
/// it on the shared loop machinery, and what makes a label on it mean what a
/// label on any other Kotlin loop means.
fn repeat_to_for_in(expr: &Expression) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "repeat") || args.len() != 2 {
        return None;
    }
    let ExprKind::Lambda { params, body, .. } = &args[1].value.kind else {
        return None;
    };
    let var = params
        .first()
        .map(|p| p.name.clone())
        .unwrap_or_else(|| "it".to_string());
    let body = match body {
        LambdaBody::Block(stmts) => stmts.clone(),
        LambdaBody::Expr(e) => vec![Statement::new(StmtKind::Expr((**e).clone()))],
    };
    Some(Statement::new(StmtKind::ForIn {
        var,
        key: None,
        iter: Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__kt_step_asc")),
            args: vec![
                Argument::positional(Expression::int(0)),
                Argument::positional(args[0].value.clone()),
                Argument::positional(Expression::int(1)),
            ],
            optional: false,
        }),
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
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
                                let base = sub.as_str().trim().to_string();
                                // `interface B : A` — B carries A's defaults, so
                                // a class implementing only B still gets them.
                                // The fold order resolves the chain, so B is
                                // augmented before any class that implements it.
                                members.push(ClassMember::Augment(AugmentDecl {
                                    from: base.clone(),
                                    via_field: None,
                                    adjustments: vec![],
                                }));
                                parents.push(base);
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
                                    if let Some(mut stmt) = walk_function_decl(inner_member) {
                                        // A Kotlin interface method with no
                                        // block is abstract; one WITH a block is
                                        // a default implementation, and the body
                                        // has to survive — `InterfaceMember`
                                        // had nowhere to put it, so every
                                        // default method was silently emptied.
                                        if let StmtKind::FunctionDecl { body, modifiers, .. } =
                                            &mut stmt.kind
                                        {
                                            modifiers.is_abstract = body.is_empty();
                                        }
                                        members.push(ClassMember::Method(Box::new(stmt)));
                                    }
                                }
                                Rule::var_decl => {
                                    if let Some(prop) = walk_class_property(inner_member.clone()) {
                                        members.push(prop);
                                        continue;
                                    }
                                    if let Some(stmt) = walk_var_decl(inner_member) {
                                        if let StmtKind::VarDecl { declarations, kind } = stmt.kind {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(pname) = decl.pattern {
                                                    members.push(ClassMember::Field {
                                                        name: pname,
                                                        type_hint: decl.type_hint,
                                                        init: decl.init,
                                                        modifiers: Modifiers {
                                                            visibility: Visibility::Public,
                                                            is_readonly: kind == VarDeclKind::Const,
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
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // An interface is a CLASS DECLARATION whose `declared_kind` says
    // `Interface` — flexclassplan §0.1's one class model, and the shape
    // `ClassKind` exists for. As a `StmtKind::InterfaceDecl` it never entered
    // `normalized_classes`, so `class W(d: I) : I by d` could not find `I`'s
    // members to promote and delegation resolved to nothing.
    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        // A Kotlin interface's supertypes are other interfaces, never a
        // superclass — so they are the interface list, not `parents`.
        interfaces: parents,
        members,
        modifiers: ClassModifiers {
            is_abstract: true,
            kind: ClassKind::Interface,
            ..Default::default()
        },
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
            Rule::destructuring_target => {
                for target_inner in inner.into_inner() {
                    if target_inner.as_rule() == Rule::identifier {
                        names.push(target_inner.as_str().to_string());
                    }
                }
            }
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
                        Rule::type_ref => type_hint = Some(type_hint_text(csub.as_str())),
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
    let mut is_operator = false;
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
                } else if m_str == "operator" {
                    is_operator = true;
                }
            }
            Rule::receiver_prefix => {
                if let Some(id_p) = inner.into_inner().next() {
                    receiver_type = Some(id_p.as_str().to_string());
                }
            }
            Rule::type_ref => {
                if return_type.is_none() && !name.is_empty() {
                    return_type = Some(type_hint_text(inner.as_str()));
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

    // `operator fun plus` is a DIFFERENT declaration from a plain `fun plus`:
    // only the former defines `+`. Kotlin's operator names are ordinary
    // identifiers, so the modifier is the only thing that distinguishes them
    // and it has to survive into `protocol.rs`, which decides slots. Encoded
    // in the name — the same device Dart uses for `operator+` — because the
    // slot mapping is a language-local decision and `Modifiers` is shared.
    // Stripped back off by `protocol::canonical_method`, so the member is
    // still stored under the name Kotlin code calls (`a.plus(b)` works).
    if is_operator && receiver_type.is_none() {
        name = format!("operator {}", name);
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
            let mut is_nullable = false;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::vararg_kw => is_rest = true,
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::type_ref => {
                        is_nullable = type_ref_is_nullable(p.as_str());
                        type_hint = Some(type_hint_text(p.as_str()));
                    }
                    Rule::expr => default = Some(walk_expr(p)),
                    _ => {}
                }
            }
            let is_optional = default.is_some();
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
    if pair.clone().into_inner().any(|p| p.as_rule() == Rule::destructuring_target) {
        return walk_destructuring_decl(pair);
    }
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
            Rule::type_ref => type_hint = Some(type_hint_text(inner.as_str())),
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    if type_hint.is_none() {
        if let Some(ref expr) = init {
            match expr.kind {
                ExprKind::Array(_) => type_hint = Some("Array".to_string()),
                ExprKind::Object(_) => type_hint = Some("Map".to_string()),
                _ => {}
            }
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

/// Is this expression a STRING by construction, so `+` on it is
/// `kotlin.String.plus` (concatenation) rather than arithmetic?
///
/// Only syntactic evidence counts — a literal, a template, or a concatenation
/// already decided. Anything requiring the operand's runtime type is left to
/// the shared path, so this never claims a `+` it cannot prove.
fn kt_is_string_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Binary { op: BinOp::Concat, .. } => true,
        // `"$a$b".trimIndent()` and friends keep the template's type.
        ExprKind::Member { object, .. } => kt_is_string_expr(object),
        _ => false,
    }
}

/// A class-body `val`/`var` that declares `get()` / `set(v)` accessors.
///
/// Kotlin's properties are not fields: `val area: Int get() = w * h` has no
/// storage at all, and `var celsius` with a custom setter must run the setter
/// on assignment. The walker used to drop the accessors on the floor and emit a
/// plain `ClassMember::Field`, so the getter never ran and the property read as
/// `undefined`. `ClassMember::Property` is the model's own shape for this —
/// C# `{ get; set; }` and Pascal properties already use it, and the compiler
/// installs `__get_`/`__set_` accessors from it.
///
/// Returns `None` for a plain stored property, which stays a field.
fn walk_class_property(pair: Pair<Rule>) -> Option<ClassMember> {
    let inners: Vec<_> = pair.into_inner().collect();
    // `var x = 1` + `private set` declares an ORDINARY stored property whose
    // setter is restricted — the accessor has no body, so there is nothing to
    // run and the backing storage is the whole implementation. Only an accessor
    // WITH a body replaces the storage; treating a bodyless one as a computed
    // property left the field unreadable (`undefined`).
    let has_accessor_body = inners.iter().any(|p| {
        p.as_rule() == Rule::property_accessor
            && p.clone()
                .into_inner()
                .any(|part| matches!(part.as_rule(), Rule::function_body_expr | Rule::block))
    });
    if !has_accessor_body {
        return None;
    }

    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;
    let mut is_readonly = false;
    let mut getter = None;
    let mut setter = None;

    for inner in inners {
        match inner.as_rule() {
            Rule::val_kw => is_readonly = true,
            Rule::identifier if name.is_empty() => name = inner.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_hint_text(inner.as_str())),
            Rule::expr => init = Some(walk_expr(inner)),
            Rule::property_accessor => {
                let mut is_get = false;
                let mut param_name = None;
                let mut body = Vec::new();
                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::get_kw => is_get = true,
                        Rule::set_kw => is_get = false,
                        Rule::identifier => param_name = Some(part.as_str().to_string()),
                        // `get() = expr` is an expression body; the model wants
                        // statements, and the value of the accessor IS its
                        // result.
                        Rule::function_body_expr => {
                            if let Some(e) = part.into_inner().next() {
                                body = vec![Statement::new(StmtKind::Return(Some(walk_expr(e))))];
                            }
                        }
                        Rule::block => body = walk_block_statements(part),
                        _ => {}
                    }
                }
                if is_get {
                    getter = Some(body);
                } else {
                    setter = Some(PropertySetter {
                        param: Param {
                            // Kotlin's implicit setter parameter is `value`.
                            name: param_name.unwrap_or_else(|| "value".to_string()),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                        body,
                    });
                }
            }
            _ => {}
        }
    }

    if name.is_empty() {
        return None;
    }
    // `var x: Int = 0  get() = field` still has backing storage; a property with
    // no initializer and a computed getter has none.
    let is_auto = init.is_some();
    Some(ClassMember::Property {
        name,
        type_hint,
        getter,
        setter,
        is_auto,
        modifiers: Modifiers {
            visibility: Visibility::Public,
            is_readonly,
            ..Default::default()
        },
    })
}

/// `this.<name>` — the read a synthesized `data class` member does.
fn this_field(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: name.to_string(),
        null_safe: false,
    })
}

/// Render a value the way Kotlin does, through the shared renderer rather than
/// by concatenating it raw. `emitter/tostring.rs` dispatches on the VALUE, so a
/// nested collection or data class renders as Kotlin spells it.
fn kt_render(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_tostring")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
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

    let mut is_data = false;
    let mut primary_prop_names = Vec::new();

    for inner in &inner_pairs {
        if inner.as_rule() == Rule::inheritance_list {
            for spec in inner.clone().into_inner() {
                if spec.as_rule() == Rule::inheritance_specifier {
                    let mut parent_name = String::new();
                    let mut spec_base_args = Vec::new();
                    let mut by_expr = None;
                    // Kotlin marks the SUPERCLASS by calling its constructor:
                    // `class D : B(n), I` extends `B` and implements `I`. An
                    // interface is never constructed, so parentheses are the
                    // whole distinction — and `B()` with no arguments has to
                    // count too, which an empty `arg_list` cannot express.
                    // Taking the FIRST supertype as the parent instead made
                    // `class C : A, B` extend `A`, so `B`'s members vanished.
                    let calls_constructor = spec.as_str().contains('(');
                    for sub in spec.into_inner() {
                        match sub.as_rule() {
                            // `type_ref` is non-atomic, so its span carries any
                            // trailing whitespace before `by` / `(`. The name is
                            // a LOOKUP KEY — `available.get("Greeter ")` misses
                            // every time — so it has to be trimmed here.
                            Rule::type_ref => parent_name = sub.as_str().trim().to_string(),
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
                        } else if calls_constructor && parents.is_empty() && !is_interface {
                            parents.push(parent_name);
                            if !spec_base_args.is_empty() {
                                base_args = Some(spec_base_args);
                            }
                        } else {
                            // An implemented interface contributes its DEFAULT
                            // methods. `interfaces` alone cannot do that — the
                            // model's own doc says that list is for identity
                            // checks and "method dispatch never walks it" — so
                            // the contribution is declared as an augmentation
                            // and the shared `class_augmentation` pass applies
                            // it, which is flexclassplan §4c's "Java interface
                            // default methods" arriving as one model rather
                            // than a fifth walker fold.
                            interfaces.push(parent_name.clone());
                            members.push(ClassMember::Augment(AugmentDecl {
                                from: parent_name,
                                via_field: None,
                                adjustments: vec![],
                            }));
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
                    "data" => is_data = true,
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
                        // `class Point(val x: Int = 0)` — the default is part of
                        // the primary constructor's signature, and `copy()`
                        // re-states it. Dropping it made every call that omitted
                        // the argument bind `undefined`.
                        let mut default = None;
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
                                Rule::type_ref => type_hint = Some(type_hint_text(p.as_str())),
                                Rule::expr => default = Some(walk_expr(p.clone())),
                                _ => {}
                            }
                        }
                        if !pname.is_empty() {
                            primary_prop_names.push(pname.clone());
                            let is_optional = default.is_some();
                            ctor_params.push(Param {
                                name: pname.clone(),
                                type_hint: type_hint.clone(),
                                default,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional,
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
                                let prop_idx = (primary_prop_names.len() - 1) as i64;
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            index: Box::new(Expression::int(prop_idx)),
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
                                    // `val area: Int get() = w * h` is a
                                    // PROPERTY, not a field — it has no storage
                                    // and the accessor has to run on each read.
                                    if let Some(prop) = walk_class_property(inner_member.clone()) {
                                        members.push(prop);
                                        continue;
                                    }
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

    if is_data {
        for (idx, pname) in primary_prop_names.iter().enumerate() {
            let comp_name = format!("component{}", idx + 1);
            members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: comp_name,
                params: vec![],
                body: vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: pname.clone(),
                    null_safe: false,
                }))))],
                return_type: None,
                modifiers: Modifiers { visibility: Visibility::Public, ..Default::default() },
                is_async: false,
                is_generator: false,
                handles: vec![],
                is_sub: false,
            }))));
        }

        if !primary_prop_names.is_empty() {
            // `Box(value=42)` — and each PART is rendered by the value renderer
            // rather than concatenated raw. A nested `data class` field prints
            // as its own `toString` this way instead of `[object]`, and a field
            // whose static type is unknown is never coerced toward a number.
            let mut str_expr = Expression::string(&format!("{}({}=", name, primary_prop_names[0]));
            str_expr = Expression::new(ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(str_expr),
                right: Box::new(kt_render(this_field(&primary_prop_names[0]))),
            });
            for pname in primary_prop_names.iter().skip(1) {
                str_expr = Expression::new(ExprKind::Binary {
                    op: BinOp::Concat,
                    left: Box::new(str_expr),
                    right: Box::new(Expression::string(&format!(", {}=", pname))),
                });
                str_expr = Expression::new(ExprKind::Binary {
                    op: BinOp::Concat,
                    left: Box::new(str_expr),
                    right: Box::new(kt_render(this_field(pname))),
                });
            }
            str_expr = Expression::new(ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(str_expr),
                right: Box::new(Expression::string(")")),
            });
            members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: "toString".to_string(),
                params: vec![],
                body: vec![Statement::new(StmtKind::Return(Some(str_expr)))],
                return_type: None,
                modifiers: Modifiers { visibility: Visibility::Public, ..Default::default() },
                is_async: false,
                is_generator: false,
                handles: vec![],
                is_sub: false,
            }))));

            let copy_params: Vec<_> = primary_prop_names.iter().map(|pname| Param {
                name: pname.clone(),
                type_hint: None,
                default: Some(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: pname.clone(),
                    null_safe: false,
                })),
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: true,
                is_nullable: false,
            }).collect();
            let copy_args: Vec<_> = primary_prop_names.iter().map(|pname| {
                Argument::positional(Expression::ident(pname))
            }).collect();
            let new_inst = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: copy_args,
                optional: false,
            });
            members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: "copy".to_string(),
                params: copy_params,
                body: vec![Statement::new(StmtKind::Return(Some(new_inst)))],
                return_type: None,
                modifiers: Modifiers { visibility: Visibility::Public, ..Default::default() },
                is_async: false,
                is_generator: false,
                handles: vec![],
                is_sub: false,
            }))));

            // `equals` / `hashCode` — Kotlin generates both for a `data class`,
            // over the PRIMARY-CONSTRUCTOR properties only. `protocol.rs` maps
            // the two spellings onto the `Eq` and `Hash` slots, so declaring
            // them as members is all it takes for `==`, `Set` membership and
            // `Map` keys to become structural. Without them a data class fell
            // back to reference identity and `P(1,2) == P(1,2)` was `false`.
            let other = "__kt_other";
            let mut eq_expr = Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident(other)),
                type_name: name.clone(),
            });
            for pname in &primary_prop_names {
                eq_expr = Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(eq_expr),
                    right: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(this_field(pname)),
                        right: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(other)),
                            field: pname.clone(),
                            null_safe: false,
                        })),
                    })),
                });
            }
            members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: "equals".to_string(),
                params: vec![Param {
                    name: other.to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: true,
                }],
                body: vec![Statement::new(StmtKind::Return(Some(eq_expr)))],
                return_type: None,
                modifiers: Modifiers { visibility: Visibility::Public, ..Default::default() },
                is_async: false,
                is_generator: false,
                handles: vec![],
                is_sub: false,
            }))));

            // Kotlin's own shape: `result = 31 * result + field.hashCode()`.
            // The per-field term is the string hash of the RENDERED field, since
            // that is the one function every value answers — and it keeps the
            // contract that matters: equal values render alike, so they hash
            // alike.
            members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: "hashCode".to_string(),
                params: vec![],
                body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_hash")),
                        args: vec![Argument::positional(kt_render(Expression::new(
                            ExprKind::This,
                        )))],
                        optional: false,
                    },
                ))))],
                return_type: None,
                modifiers: Modifiers { visibility: Visibility::Public, ..Default::default() },
                is_async: false,
                is_generator: false,
                handles: vec![],
                is_sub: false,
            }))));
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
            // What this declaration DECLARES. `emit_class_from_ast` copies it
            // into `NormalClass.declared_kind` and the compiler stamps it on
            // the class object — which is what answers `interface_exists` and
            // keeps an interface from being treated as an instantiable class.
            // Left at the `Class` default, every Kotlin `interface` claimed to
            // be one.
            kind: if is_interface {
                ClassKind::Interface
            } else {
                ClassKind::Class
            },
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
        for (idx, name) in destruct_names.clone().into_iter().enumerate() {
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

    let final_iter = if !destruct_names.is_empty() {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__coll_is_array")),
                args: vec![Argument::positional(iter_expr.clone())],
                optional: false,
            })),
            then: Box::new(iter_expr.clone()),
            else_: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dict_items")),
                args: vec![Argument::positional(iter_expr)],
                optional: false,
            })),
        })
    } else {
        iter_expr
    };

    Some(Statement::new(StmtKind::ForIn {
        var: var_id,
        key: None,
        iter: final_iter,
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
    let mut prefix_stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::lambda_params => {
                for lp in inner.into_inner() {
                    if lp.as_rule() == Rule::lambda_param {
                        let mut name = String::new();
                        let mut destruct_names = Vec::new();
                        let mut type_hint = None;
                        let mut lambda_param_nullable = false;
                        for lsub in lp.into_inner() {
                            match lsub.as_rule() {
                                Rule::identifier => name = lsub.as_str().to_string(),
                                Rule::lambda_destructure => {
                                    for sub in lsub.into_inner() {
                                        if sub.as_rule() == Rule::identifier {
                                            destruct_names.push(sub.as_str().to_string());
                                        }
                                    }
                                }
                                Rule::type_ref => {
                                    lambda_param_nullable = type_ref_is_nullable(lsub.as_str());
                                    type_hint = Some(type_hint_text(lsub.as_str()));
                                }
                                _ => {}
                            }
                        }
                        if !destruct_names.is_empty() {
                            let tmp_param = gen_tmp_name();
                            params.push(Param {
                                name: tmp_param.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                            for (idx, dname) in destruct_names.into_iter().enumerate() {
                                prefix_stmts.push(Statement::new(StmtKind::VarDecl {
                                    declarations: vec![VarDeclarator {
                                        pattern: BindingPattern::Ident(dname),
                                        type_hint: None,
                                        init: Some(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::ident(&tmp_param)),
                                            index: Box::new(Expression::int(idx as i64)),
                                            null_safe: false,
                                        })),
                                        array_bounds: None,
                                        with_events: false,
                                    }],
                                    kind: VarDeclKind::Const,
                                }));
                            }
                        } else if !name.is_empty() {
                            params.push(Param {
                                name,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: lambda_param_nullable,
                            });
                        }
                    }
                }
                // prefix_stmts stored in outer scope
            }
            Rule::statement => {
                if let Some(s) = walk_statement(inner) {
                    body.push(s);
                }
            }
            _ => {}
        }
    }

    if !prefix_stmts.is_empty() {
        prefix_stmts.extend(body);
        body = prefix_stmts;
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
                        "+=" => {
                            if matches!(rhs.kind, ExprKind::Binary { op: BinOp::Concat, .. } | ExprKind::Lit(Literal::Str(_)))
                                || matches!(lhs.kind, ExprKind::Lit(Literal::Str(_)))
                            {
                                BinOp::Concat
                            } else {
                                BinOp::Add
                            }
                        }
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
                    // `"a" + x` resolves to `kotlin.String.plus` — CONCATENATION
                    // for every right operand, whatever its type. Emitting a
                    // generic `Add` left the decision to whatever the shared
                    // type inference could see, and an operand it could not
                    // classify (a member read on a user object, a call result)
                    // was coerced toward a number: `"x=" + this.n` trapped in
                    // `toF64` even though `n` held a string. Left-associativity
                    // carries the answer along a chain, so testing the LEFT
                    // operand covers `"a" + x + y`.
                    "+" if kt_is_string_expr(&current) => BinOp::Concat,
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
                let op_str = op_pair.as_str();
                // `/` used to be rewritten to `(a / b) | 0` here — the JS
                // integer-truncation idiom, applied UNCONDITIONALLY. Kotlin
                // truncates only when BOTH operands are integers, so this made
                // `7.0 / 2.0` answer 3. The shared emitter decides now, from
                // `integer_division_on_slash` plus this language's
                // `[builtin_types] int` spellings (builtinslotplan.md §3i).
                let op = match op_str {
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
                    // Kotlin `a to b` → Pair(a, b)
                    current = create_pair_expr(current, next_expr);
                } else if op_str == "until" {
                    // a until b → exclusive ascending [a, a+1, ..., b-1]
                    current = Expression::new(ExprKind::Range {
                        start: Box::new(current),
                        end: Box::new(next_expr),
                        inclusive: false,
                    });
                } else if op_str == "downTo" {
                    // a downTo b → descending [a, a-1, ..., b], INCLUSIVE of b.
                    // Maps to `__kt_step_desc(a, b - 1, -1)` → `collections.range_step`,
                    // whose stop is EXCLUSIVE (it iterates while `i > stop` for a
                    // negative step). Passing `b` straight through dropped the last
                    // element: `5 downTo 2` yielded 5,4,3.
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_step_desc")),
                        args: vec![
                            Argument::positional(current),
                            Argument::positional(Expression::new(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(next_expr),
                                right: Box::new(Expression::int(1)),
                            })),
                            Argument::positional(Expression::new(ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr: Box::new(Expression::int(1)),
                            })),
                        ],
                        optional: false,
                    });
                } else if op_str == "step" {
                    // `range step n` — must convert range to 3-arg stepped form.
                    match current.kind.clone() {
                        // (a downTo b) step n  → replace -1 with -n
                        ExprKind::Call { callee, mut args, optional }
                            if matches!(&callee.kind, ExprKind::Ident(nm) if nm == "__kt_step_desc") =>
                        {
                            if args.len() == 3 {
                                args[2] = Argument::positional(Expression::new(ExprKind::Unary {
                                    op: UnaryOp::Neg,
                                    expr: Box::new(next_expr),
                                }));
                            }
                            current = Expression::new(ExprKind::Call { callee, args, optional });
                        }
                        // (a..b) step n  or  (a until b) step n
                        ExprKind::Range { start, end, inclusive } => {
                            let stop = if inclusive {
                                // inclusive end+1 so the 3-arg exclusive loop includes end
                                Expression::new(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: end,
                                    right: Box::new(Expression::int(1)),
                                })
                            } else {
                                *end
                            };
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_step_asc")),
                                args: vec![
                                    Argument::positional(*start),
                                    Argument::positional(stop),
                                    Argument::positional(next_expr),
                                ],
                                optional: false,
                            });
                        }
                        _ => {
                            // Fallback: pass through as method call
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(current),
                                    field: "step".to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional(next_expr)],
                                optional: false,
                            });
                        }
                    }
                } else if let Some(op) = infix_bitwise_op(op_str) {
                    // Kotlin spells the bitwise operators as infix functions
                    // (`6 and 3`, `1 shl 2`). They are the SAME operators every
                    // other language writes with punctuation, so they lower to
                    // the shared `BinOp` and reach `primitives/operators.rs`
                    // rather than becoming an `Int.and(…)` member call that no
                    // primitive implements.
                    current = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(current),
                        right: Box::new(next_expr),
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
                    Rule::type_args => {
                        continue;
                    }
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
                                                Rule::identifier => {
                                                    if arg_name.is_none() && arg_expr.is_none() {
                                                        arg_name = Some(sub.as_str().to_string());
                                                    }
                                                }
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

                        if let ExprKind::Member { ref object, ref field, null_safe: _ } = current.clone().kind {
                            match field.as_str() {
                                "put" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        value: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        right: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "containsKey" | "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::In,
                                        left: Box::new(args[0].value.clone()),
                                        right: object.clone(),
                                    });
                                    continue;
                                }
                                // NOTE: `.add(x)` for Set (dict) is handled in the second
                                // Member block below via __coll_push, which works uniformly
                                // for both list (array.push) and set (set semantics via
                                // array.push on the keys array). Do NOT intercept it here.

                                "remove" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Delete(Box::new(Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    }))));
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    let keys_len = Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: Box::new(keys_len),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    let keys_len = Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(keys_len),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }

                        let func_name = match &current.kind {
                            ExprKind::Ident(name) => Some(name.clone()),
                            _ => None,
                        };

                        if let Some(ref fn_name) = func_name {
                            if fn_name == "Pair" && args.len() == 2 {
                                current = create_pair_expr(args[0].value.clone(), args[1].value.clone());
                                continue;
                            }
                            if fn_name == "Triple" && args.len() == 3 {
                                current = create_triple_expr(args[0].value.clone(), args[1].value.clone(), args[2].value.clone());
                                continue;
                            }
                            if matches!(fn_name.as_str(), "mapOf" | "mutableMapOf" | "linkedMapOf" | "hashMapOf" | "buildMap" | "emptyMap") {
                                let mut props = Vec::new();
                                for arg in args {
                                    if let ExprKind::Object(ref pair_props) = arg.value.kind {
                                        let mut k_expr = None;
                                        let mut v_expr = None;
                                        for p in pair_props {
                                            if let ObjectProperty::KeyValue { key, value } = p {
                                                if let ExprKind::Lit(Literal::Str(s)) = &key.kind {
                                                    if s == "0" || s == "first" {
                                                        if k_expr.is_none() { k_expr = Some(value.clone()); }
                                                    } else if s == "1" || s == "second" {
                                                        if v_expr.is_none() { v_expr = Some(value.clone()); }
                                                    }
                                                }
                                            }
                                        }
                                        if let (Some(k), Some(v)) = (k_expr, v_expr) {
                                            props.push(ObjectProperty::KeyValue { key: k, value: v });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Tuple(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].clone(),
                                                value: pair_elems[1].clone(),
                                            });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Array(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].value.clone(),
                                                value: pair_elems[1].value.clone(),
                                            });
                                            continue;
                                        }
                                    }
                                    props.push(ObjectProperty::KeyValue {
                                        key: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(0)),
                                            null_safe: false,
                                        }),
                                        value: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(1)),
                                            null_safe: false,
                                        }),
                                    });
                                }
                                current = create_map_expr(props);
                                continue;
                            }
                            if matches!(fn_name.as_str(), "setOf" | "mutableSetOf" | "linkedSetOf" | "hashSetOf" | "buildSet" | "emptySet") {
                                let elems = args.into_iter().map(|a| a.value).collect();
                                current = create_kotlin_set_expr(elems);
                                continue;
                            }
                            if matches!(fn_name.as_str(), "listOf" | "mutableListOf" | "arrayOf" | "emptyList" | "intArrayOf" | "doubleArrayOf" | "booleanArrayOf" | "charArrayOf" | "longArrayOf" | "sequenceOf") {
                                let elements = args.into_iter().map(|a| ArrayElement {
                                    key: None,
                                    value: a.value,
                                    spread: false,
                                    by_ref: false,
                                }).collect();
                                current = Expression::new(ExprKind::Array(elements));
                                continue;
                            }
                        }

                        if let ExprKind::Member { ref object, ref field, .. } = current.kind {
                            match field.as_str() {
                                "put" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        value: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    let get_expr = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(get_expr),
                                        right: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "containsKey" | "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_has")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "containsValue" if args.len() == 1 => {
                                    let values_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_contains")),
                                        args: vec![
                                            Argument::positional(values_expr),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    let sz = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_size")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: Box::new(sz),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    let sz = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_size")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(sz),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "remove" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Delete(Box::new(Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    }))));
                                    continue;
                                }
                                "removeAt" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_removeAt")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "add" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_push")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "add" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_insert")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(args[1].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "indexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_indexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "lastIndexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_lastIndexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "reversed" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_reverse")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sorted" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sorted")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "joinToString" if !args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "joinToString" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(", ".to_string())))),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sum" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sum")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "fold" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: "__array_reduce".to_string(),
                                            null_safe: false,
                                        })),
                                        args: vec![
                                            Argument::positional(args[1].value.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "take" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(Expression::int(0)),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "drop" if args.len() == 1 => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(len_expr),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "first" | "firstOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "max" | "maxOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_max")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "min" | "minOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_min")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "filter" | "filterNot" | "map" | "forEach" if args.len() == 1 => {
                                    let iter_target = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__coll_is_array")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                        then: Box::new(*object.clone()),
                                        else_: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__dict_items")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                    });
                                    // Keep the SOURCE spelling. `[array_methods]`
                                    // is keyed by it (`filter = "__array_filter"`),
                                    // and `lookup_array_method` looks up the KEY —
                                    // rewriting the field to `__array_filter` here
                                    // meant the lookup missed, the higher-order
                                    // dispatch never ran, and the call fell through
                                    // to a member named `__array_filter` that does
                                    // not exist ("undefined is not callable").
                                    let emit_method = match field.as_str() {
                                        "filterNot" => "filter",
                                        other => other,
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(iter_target),
                                            field: emit_method.to_string(),
                                            null_safe: false,
                                        })),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }

                        // Kotlin has no `new`, so a call whose callee names a
                        // TYPE is a construction. That is true of a qualified
                        // spelling too — `java.util.ArrayList()`,
                        // `java.math.BigInteger("1")` — which stayed an
                        // ordinary member call and trapped, while the
                        // `import`ed form worked. Same rule, applied to the
                        // last segment of the chain.
                        let is_type_spelling = |name: &str| {
                            name.chars().next().is_some_and(char::is_uppercase)
                                && !matches!(
                                    name,
                                    "Exception"
                                        | "IllegalArgumentException"
                                        | "IllegalStateException"
                                        | "NullPointerException"
                                        | "IndexOutOfBoundsException"
                                )
                        };
                        let is_class_name = match &current.kind {
                            ExprKind::Ident(name) => is_type_spelling(name),
                            // Only a chain of plain idents is a qualified type
                            // name; `expr.Foo()` is a method call on a value.
                            ExprKind::Member { object, field, .. } => {
                                is_type_spelling(field) && is_ident_chain(object)
                            }
                            _ => false,
                        };

                        if is_class_name {
                            current = Expression::new(ExprKind::New {
                                class: Box::new(current),
                                args,
                            });
                        } else {
                            if let ExprKind::Member { ref object, ref field, .. } = current.kind {
                                if matches!(field.as_str(), "filter" | "filterNot" | "map" | "forEach") && args.len() == 1 {
                                    let iter_target = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__coll_is_array")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                        then: Box::new(*object.clone()),
                                        else_: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__dict_items")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                    });
                                    // Keep the SOURCE spelling. `[array_methods]`
                                    // is keyed by it (`filter = "__array_filter"`),
                                    // and `lookup_array_method` looks up the KEY —
                                    // rewriting the field to `__array_filter` here
                                    // meant the lookup missed, the higher-order
                                    // dispatch never ran, and the call fell through
                                    // to a member named `__array_filter` that does
                                    // not exist ("undefined is not callable").
                                    let emit_method = match field.as_str() {
                                        "filterNot" => "filter",
                                        other => other,
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(iter_target),
                                            field: emit_method.to_string(),
                                            null_safe: false,
                                        })),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                            }
                            // `super.f(a)` — `member_suffix` already turned
                            // `super.f` into a `SuperCall`, which IS the call.
                            // Wrapping it again called the RESULT ("string is
                            // not callable"), so the arguments land on the node
                            // that already exists instead of a second one.
                            if let ExprKind::SuperCall {
                                args: super_args, ..
                            } = &mut current.kind
                            {
                                if super_args.is_empty() {
                                    *super_args = args;
                                    continue;
                                }
                            }
                            // `.toString` used to be rewritten into `"" + x`
                            // here, so the `()` that follows it had to be
                            // swallowed. That rewrite is gone: `toString` is a
                            // MEMBER, a class may override it, and the built-in
                            // rendering is declared as a value method
                            // (`common:kotlin.tostring`) instead. Rendering
                            // dispatches on the VALUE, in `emitter/tostring.rs`.
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
                                // `first`/`second`/`third` are PROPERTIES on
                                // `Pair`/`Triple`, which lower to an array
                                // (`common:collections.new`), so the positional
                                // read is the whole meaning. `componentN` is
                                // deliberately NOT here: it is a FUNCTION, and a
                                // `data class` declares its own — rewriting the
                                // member turned `u.component1()` into `u[0]()`
                                // ("string is not callable") and made the
                                // synthesized member unreachable. It resolves as
                                // an ordinary member call now, like any other.
                                "first" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false,
                                    });
                                }
                                "second" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(1)),
                                        null_safe: false,
                                    });
                                }
                                "third" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(2)),
                                        null_safe: false,
                                    });
                                }
                                "keys" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_keys")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "values" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "entries" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_items")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "size" | "length" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "lastIndex" => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Sub,
                                        left: Box::new(len_expr),
                                        right: Box::new(Expression::int(1)),
                                    });
                                }
                                "indices" => {
                                    let len_expr = Expression::new(ExprKind::Member {
                                        object: Box::new(current.clone()),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Range {
                                        start: Box::new(Expression::int(0)),
                                        end: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Sub,
                                            left: Box::new(len_expr),
                                            right: Box::new(Expression::int(1)),
                                        })),
                                        inclusive: true,
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
                    Rule::inc_suffix => {
                        let op_str = suffix_inner.as_str();
                        let bin_op = if op_str == "++" { BinOp::Add } else { BinOp::Sub };
                        current = Expression::new(ExprKind::Assign {
                            target: Box::new(current.clone()),
                            value: Box::new(Expression::new(ExprKind::Binary {
                                op: bin_op,
                                left: Box::new(current),
                                right: Box::new(Expression::int(1)),
                            })),
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
                // `this@Outer` / `super<Base>` / `super@Outer`. The label and
                // the explicit supertype are RESOLUTION hints; the receiver
                // itself is still `this` / `super`, so the concept node is the
                // same one an unqualified occurrence produces and no downstream
                // path has to learn a second shape.
                Rule::this_expr => Expression::new(ExprKind::This),
                Rule::super_expr => Expression::new(ExprKind::Super),
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

/// The type a `type_ref` names, with Kotlin's nullability marker removed.
///
/// `String?` and `String` are ONE type as far as the shared machinery is
/// concerned — nullability is carried by `Param::is_nullable`, not by the
/// hint's spelling. Stripping it here is what lets `[builtin_types]` declare
/// each spelling once (`builtinslotplan.md` step 4a); leaving the `?` on would
/// make every declared spelling need a nullable twin, and `String?` would
/// resolve to no built-in at all.
fn type_hint_text(raw: &str) -> String {
    raw.trim().trim_end_matches('?').trim_end().to_string()
}

/// Whether a `type_ref`'s source text carries Kotlin's `?`.
fn type_ref_is_nullable(raw: &str) -> bool {
    raw.trim_end().ends_with('?')
}

/// Kotlin's infix spellings of the bitwise operators -> the shared `BinOp`.
///
/// `ushr` has no `BinOp` of its own; Kotlin's other shift is arithmetic, so
/// `shr` takes `Shr` and `ushr` stays a member call until an unsigned shift
/// exists to route it to.
fn infix_bitwise_op(op: &str) -> Option<BinOp> {
    match op {
        "and" => Some(BinOp::BitAnd),
        "or" => Some(BinOp::BitOr),
        "xor" => Some(BinOp::BitXor),
        "shl" => Some(BinOp::Shl),
        "shr" => Some(BinOp::Shr),
        _ => None,
    }
}

/// `0xFF` / `0b1010` / `1_000_000` / `12L` / `7u` -> the integer value.
///
/// The `_` grouping is a lexical convenience with no value, and the `u`/`L`
/// suffixes only pick the Kotlin static type; both are stripped before the
/// radix parse so one function covers every spelling.
fn parse_int_literal(raw: &str) -> i64 {
    let cleaned: String = raw.chars().filter(|c| *c != '_').collect();
    let body = cleaned.trim_end_matches(['L', 'l', 'u', 'U']);
    let (digits, radix) = if let Some(rest) = body.strip_prefix("0x").or_else(|| body.strip_prefix("0X")) {
        (rest, 16)
    } else if let Some(rest) = body.strip_prefix("0b").or_else(|| body.strip_prefix("0B")) {
        (rest, 2)
    } else {
        (body, 10)
    };
    i64::from_str_radix(digits, radix).unwrap_or(0)
}

fn walk_literal(pair: Pair<Rule>) -> Expression {
    let inner = pair.into_inner().next().unwrap();
    match inner.as_rule() {
        Rule::null_kw => Expression::null(),
        Rule::true_kw => Expression::bool(true),
        Rule::false_kw => Expression::bool(false),
        Rule::int_literal => Expression::int(parse_int_literal(inner.as_str())),
        Rule::float_literal => {
            let s: String = inner
                .as_str()
                .chars()
                .filter(|c| *c != '_')
                .collect::<String>()
                .trim_end_matches(['f', 'F'])
                .to_string();
            Expression::float(s.parse::<f64>().unwrap_or(0.0))
        }
        Rule::string_literal => walk_string_literal(inner),
        Rule::char_literal => {
            let s = inner.as_str();
            let content = &s[1..s.len().saturating_sub(1)];
            let decoded = match content {
                "\\n" => "\n".to_string(),
                "\\t" => "\t".to_string(),
                "\\r" => "\r".to_string(),
                "\\\"" => "\"".to_string(),
                "\\'" => "'".to_string(),
                "\\\\" => "\\".to_string(),
                "\\$" => "$".to_string(),
                "\\b" => "\x08".to_string(),
                "\\f" => "\x0C".to_string(),
                s if s.starts_with("\\u") && s.len() == 6 => {
                    if let Ok(code) = u32::from_str_radix(&s[2..], 16) {
                        if let Some(ch) = char::from_u32(code) {
                            ch.to_string()
                        } else {
                            s.to_string()
                        }
                    } else {
                        s.to_string()
                    }
                }
                s => s.to_string(),
            };
            Expression::string(&decoded)
        }
        _ => Expression::null(),
    }
}

fn walk_string_literal(pair: Pair<Rule>) -> Expression {
    let mut parts = Vec::new();
    collect_string_parts(pair, &mut parts);

    // `str_text` matches ONE character, so `"x="` arrives as two parts and a
    // plain literal was never a `Lit(Str)` at all — it was a `Binary` tree one
    // node per character. Everything downstream that asks "is this a string?"
    // (the `+` decision below, `[builtin_types]` classification, constant
    // folding) answered no for every literal longer than one char. Fold the
    // adjacent literal runs back into the single literal the source wrote.
    let mut folded: Vec<Expression> = Vec::new();
    for part in parts {
        match (&part.kind, folded.last_mut().map(|last| &mut last.kind)) {
            (ExprKind::Lit(Literal::Str(text)), Some(ExprKind::Lit(Literal::Str(acc)))) => {
                acc.push_str(text);
            }
            _ => folded.push(part),
        }
    }

    if folded.is_empty() {
        Expression::string("")
    } else if folded.len() == 1 {
        folded.remove(0)
    } else {
        // A template IS concatenation — `"a${x}b"` never means arithmetic, and
        // each interpolated part is already rendered by `__kt_tostring`.
        let mut iter = folded.into_iter();
        let mut acc = iter.next().unwrap();
        for p in iter {
            acc = Expression::new(ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(acc),
                right: Box::new(p),
            });
        }
        acc
    }
}

/// One `$x` / `${expr}` inside a string template.
///
/// Kotlin templates call `toString()` on the part — they do NOT lean on `+`
/// coercion, and the two disagree: a `Boolean` concatenates as `1`, a `List` as
/// `1,2,3`, a `Map` as `[object]`. Routing the part through the same renderer
/// `println` uses is what makes `"v=$flag"` and `println(flag)` agree, which is
/// the whole point of having one renderer.
fn interpolated_part(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_tostring")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn collect_string_parts(pair: Pair<Rule>, parts: &mut Vec<Expression>) {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::raw_string_literal | Rule::plain_string_literal | Rule::string_content | Rule::raw_string_content => {
                collect_string_parts(child, parts);
            }
            Rule::str_text | Rule::raw_str_text => {
                parts.push(Expression::string(child.as_str()));
            }
            Rule::str_escaped => {
                let s = child.as_str();
                let decoded = match s {
                    "\\n" => "\n".to_string(),
                    "\\t" => "\t".to_string(),
                    "\\r" => "\r".to_string(),
                    "\\\"" => "\"".to_string(),
                    "\\'" => "'".to_string(),
                    "\\\\" => "\\".to_string(),
                    "\\$" => "$".to_string(),
                    "\\b" => "\x08".to_string(),
                    "\\f" => "\x0C".to_string(),
                    s if s.starts_with("\\u") && s.len() == 6 => {
                        if let Ok(code) = u32::from_str_radix(&s[2..], 16) {
                            if let Some(ch) = char::from_u32(code) {
                                ch.to_string()
                            } else {
                                s.to_string()
                            }
                        } else {
                            s.to_string()
                        }
                    }
                    _ => s.to_string(),
                };
                parts.push(Expression::string(&decoded));
            }
            Rule::str_interpolated_var => {
                if let Some(id_pair) = child.into_inner().next() {
                    parts.push(interpolated_part(Expression::ident(id_pair.as_str())));
                }
            }
            Rule::str_interpolated_expr => {
                let raw = child.as_str();
                if raw.starts_with("${") && raw.ends_with('}') {
                    let inner_str = raw[2..raw.len()-1].trim();
                    let unescaped = inner_str.replace("\\\"", "\"");
                    if let Ok(mut pairs) = KotlinParser::parse(Rule::expr, &unescaped) {
                        if let Some(epair) = pairs.next() {
                            parts.push(interpolated_part(walk_expr(epair)));
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

/// `a to b` / `Pair(a, b)`.
///
/// `ExprKind::Tuple`, not `ExprKind::Array`: a Pair and a two-element List are
/// the same runtime array, and only the tuple TAG (`tuple_literals_tagged` in
/// the profile) tells them apart — which is what lets `println` render one as
/// `(3, x)` and the other as `[3, x]`.
fn create_pair_expr(a: Expression, b: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![a, b]))
}

/// `Triple(a, b, c)` — see [`create_pair_expr`].
fn create_triple_expr(a: Expression, b: Expression, c: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![a, b, c]))
}

/// `mapOf(…)` / a Kotlin map literal.
///
/// Plain data. This used to append a synthesised `toString` PROPERTY to the
/// object, which made the map render itself — and put `toString` into the
/// map's own `__keys`, so the map contained a member the program never put
/// there (`{a=1, b=2, toString=…}` once the renderer stopped hiding it).
/// Rendering is `emitter/tostring.rs`'s job; a map is just its entries.
fn create_map_expr(props: Vec<ObjectProperty>) -> Expression {
    Expression::new(ExprKind::Object(props))
}

/// `setOf(…)` / `mutableSetOf(…)`.
///
/// A Kotlin `Set` is a dict whose values are all `true` — the keys ARE the
/// elements, which is what gives `in` its O(1) answer. It carries
/// [`SET_MARKER`] because a `Set` and a `Map` are the same runtime shape and
/// render differently: `[1, 2, 3]` versus `{a=1}`.
fn create_kotlin_set_expr(elems: Vec<Expression>) -> Expression {
    let mut props = Vec::with_capacity(elems.len() + 1);
    props.push(ObjectProperty::KeyValue {
        key: Expression::string(SET_MARKER),
        value: Expression::bool(true),
    });
    for elem in elems {
        props.push(ObjectProperty::KeyValue {
            key: elem,
            value: Expression::bool(true),
        });
    }
    Expression::new(ExprKind::Object(props))
}
