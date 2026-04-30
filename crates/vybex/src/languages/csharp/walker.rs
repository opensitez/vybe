use pest::Parser;
use pest::iterators::Pair;
use super::{CSharpParser, Rule};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = CSharpParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => { body.push(walk_statement(top)?); continue; }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI => continue,
                Rule::using_directive => imports.push(walk_using(pair)?),
                Rule::namespace_declaration => {
                    // Build NamespaceDecl with name and body
                    let mut ns_name = String::new();
                    let mut ns_body: Vec<Statement> = Vec::new();
                    for p in pair.into_inner() {
                        match p.as_rule() {
                            Rule::dotted_name => ns_name = p.as_str().to_string(),
                            Rule::using_directive => imports.push(walk_using(p)?),
                            _ => {
                                if let Ok(stmt) = walk_top_level(p) {
                                    ns_body.push(stmt);
                                }
                            }
                        }
                    }
                    body.push(Statement::new(StmtKind::NamespaceDecl {
                        name: ns_name,
                        body: ns_body,
                    }));
                }
                _ => {
                    if let Ok(stmt) = walk_top_level(pair) {
                        body.push(stmt);
                    }
                }
            }
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::CSharp,
        body,
        imports,
    })
}

// ── Top-level items ─────────────────────────────────────────────────────────

fn walk_top_level(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::struct_declaration => walk_struct_decl(pair)?,
        Rule::interface_declaration => walk_interface_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::record_declaration => walk_record_decl(pair)?,
        Rule::delegate_declaration => walk_delegate_decl(pair)?,
        _ => walk_statement(pair)?.kind,
    };
    Ok(Statement::with_span(kind, span))
}

// ── Statements ──────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => {
            let stmts = pair.into_inner()
                .map(walk_statement)
                .collect::<Result<Vec<_>, _>>()?;
            StmtKind::Block(stmts)
        }
        Rule::local_var_declaration => walk_local_var(pair)?,
        Rule::tuple_deconstruction_decl => walk_tuple_deconstruction(pair)?,
        Rule::if_statement => walk_if(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::foreach_statement => walk_foreach(pair)?,
        Rule::while_statement => walk_while(pair)?,
        Rule::do_while_statement => walk_do_while(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::yield_statement => walk_yield_stmt(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::throw_statement => walk_throw(pair)?,
        Rule::try_statement => walk_try(pair)?,
        Rule::using_statement => walk_using_stmt(pair)?,
        Rule::lock_statement => walk_lock(pair)?,
        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().ok_or("Empty expr stmt")?)?;
            // Check if this is a compound assignment or event subscription
            classify_expr_stmt(expr)
        }
        // Type declarations can appear inside methods
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::struct_declaration => walk_struct_decl(pair)?,
        Rule::interface_declaration => walk_interface_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::record_declaration => walk_record_decl(pair)?,
        Rule::delegate_declaration => walk_delegate_decl(pair)?,
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

/// Classify an expression statement — detect assignment, compound assignment,
/// and event `+=` / `-=` patterns. Event subscription becomes the canonical
/// `AddHandler` / `RemoveHandler` AST node so the compiler routes it through
/// `compiler_common::gui::emit_bind_event` (the same path VB `Handles`,
/// JS `addEventListener`, Python `bind`, etc. resolve to).
fn classify_expr_stmt(expr: Expression) -> StmtKind {
    // Detect `obj.Event += handler` (compound add on a member access). The
    // walker rewrote `+=` as `target = target + value`, so the shape is:
    //   Assign { target: Member { obj, "Event" },
    //            value: Binary { Add, Member { obj, "Event" }, handler } }
    //
    // We only treat this as event subscription when the right-hand side is
    // function-shaped (an identifier, member access, or lambda). A literal,
    // arithmetic expression, etc. means the user is doing genuine numeric
    // compound assignment on a property and we leave it alone.
    if let ExprKind::Assign { target, value } = &expr.kind {
        if let ExprKind::Member { object: ev_obj, field: ev_field, .. } = &target.kind {
            if let ExprKind::Binary { op, left, right } = &value.kind {
                let same_target = matches!(&left.kind, ExprKind::Member { object, field, .. }
                    if member_eq(object, field, ev_obj, ev_field));
                let handler_shape = matches!(&right.kind,
                    ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Lambda { .. }
                );
                if same_target && handler_shape && (matches!(op, BinOp::Add) || matches!(op, BinOp::Sub)) {
                    let event_name = ev_field.to_lowercase();
                    let control = (**ev_obj).clone();
                    let handler = (**right).clone();
                    return if matches!(op, BinOp::Add) {
                        StmtKind::AddHandler { control, event: event_name, handler }
                    } else {
                        StmtKind::RemoveHandler { control, event: event_name, handler }
                    };
                }
            }
        }
    }

    match expr.kind {
        ExprKind::Assign { target, value } => {
            // Check for compound assignment patterns
            StmtKind::Assign { targets: vec![*target], value: *value }
        }
        _ => StmtKind::Expr(expr),
    }
}

/// Compare two member access targets for structural equality. Used to detect
/// the `+=`/`-=` shape where the LHS and the LHS-of-the-compound-binary are
/// the same control event (e.g. `btn.Click += h` desugared to
/// `btn.Click = btn.Click + h`).
fn member_eq(obj_a: &Expression, field_a: &str, obj_b: &Expression, field_b: &str) -> bool {
    if !field_a.eq_ignore_ascii_case(field_b) { return false; }
    match (&obj_a.kind, &obj_b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (ExprKind::This, ExprKind::This) => true,
        _ => false,
    }
}

// ── Using directive ─────────────────────────────────────────────────────────

fn walk_using(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let path = pair.into_inner()
        .find(|p| p.as_rule() == Rule::dotted_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();
    Ok(Import {
        kind: ImportKind::Simple { path, alias: None },
        span,
    })
}

// ── Variable declaration ────────────────────────────────────────────────────

/// `var (a, b) = f();` → `Assign { targets: [Destructure(Array([a, b]))], value: f() }`.
/// The compiler's multi-value receive path auto-defines any unresolved
/// idents as new locals inside a function, so a single `Assign` statement
/// is enough — no separate `VarDecl` is needed.
fn walk_tuple_deconstruction(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut idents: Vec<String> = Vec::new();
    let mut value: Option<Expression> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw => {}
            Rule::ident_name => idents.push(p.as_str().to_string()),
            Rule::expression => {
                value = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }
    let value = value.ok_or("tuple deconstruction missing RHS")?;
    let patterns: Vec<ArrayPatternElem> = idents.into_iter()
        .map(|n| ArrayPatternElem::Pattern(BindingPattern::Ident(n), None))
        .collect();
    let target = Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)));
    Ok(StmtKind::Assign { targets: vec![target], value })
}

fn walk_local_var(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty local var")?;

    // Skip type name (var or explicit type)
    let type_hint = match first.as_rule() {
        Rule::var_kw => None,
        Rule::type_name => Some(first.as_str().to_string()),
        _ => None,
    };

    let mut declarations = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::var_declarator_list => {
                for vd in p.into_inner() {
                    if vd.as_rule() == Rule::var_declarator {
                        declarations.push(walk_var_declarator(vd)?);
                    }
                }
            }
            Rule::var_declarator => declarations.push(walk_var_declarator(p)?),
            _ => {}
        }
    }

    if let Some(type_hint) = type_hint {
        for decl in &mut declarations {
            decl.type_hint = Some(type_hint.clone());
        }
    } else {
        // `var` type inference: `var x = new ClassName(...)` infers
        // type=ClassName so Component Model instance-method dispatch
        // (`x.Method(...)`) can resolve at compile time. Without this,
        // typed-receiver dispatch falls through to runtime hint lookup
        // via `__type` stamping — slower and weaker. Handles both
        // bare names (`new Dictionary()`) and namespace-qualified
        // names (`new System.Text.StringBuilder()`) — the last segment
        // of the dotted class is the unqualified class name.
        for decl in &mut declarations {
            if decl.type_hint.is_none() {
                if let Some(ref init) = decl.init {
                    if let crate::ast::ExprKind::New { class, .. } = &init.kind {
                        let inferred = match &class.kind {
                            crate::ast::ExprKind::Ident(n) => Some(n.clone()),
                            crate::ast::ExprKind::Member { field, .. } => Some(field.clone()),
                            _ => None,
                        };
                        if let Some(name) = inferred {
                            decl.type_hint = Some(name);
                        }
                    }
                }
            }
        }
    }

    Ok(StmtKind::VarDecl { declarations, kind: VarDeclKind::Let })
}

fn walk_var_declarator(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut inner = pair.into_inner();
    let name = inner.next().ok_or("Empty var declarator")?.as_str().to_string();
    let init = inner.next().map(walk_expression).transpose()?;
    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: None,
        init,
        array_bounds: None,
        with_events: false,
    })
}

// ── Class declaration ───────────────────────────────────────────────────────

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();
    let mut class_mods = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        "public" => class_mods.visibility = Visibility::Public,
                        "private" => class_mods.visibility = Visibility::Private,
                        "protected" => class_mods.visibility = Visibility::Protected,
                        "internal" => class_mods.visibility = Visibility::Internal,
                        s if s.starts_with("static") => class_mods.is_static = true,
                        s if s.starts_with("abstract") => class_mods.is_abstract = true,
                        s if s.starts_with("sealed") => class_mods.is_sealed = true,
                        s if s.starts_with("partial") => class_mods.is_partial = true,
                        _ => {}
                    }
                }
            }
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::base_list => {
                let mut first = true;
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        let type_str = bp.as_str().trim().to_string();
                        // First is parent class, rest are interfaces (heuristic: starts with I)
                        if first {
                            // Check if it's a known framework type or starts with uppercase (not I)
                            if type_str.starts_with('I') && type_str.len() > 1 && type_str.chars().nth(1).map_or(false, |c| c.is_uppercase()) {
                                interfaces.push(type_str);
                            } else {
                                parents.push(type_str);
                            }
                            first = false;
                        } else {
                            interfaces.push(type_str);
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.push(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl { name, parents, interfaces, members, modifiers: class_mods })
}

fn walk_class_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut mods = Modifiers::default();
    let mut member_pair = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        "public" => mods.visibility = Visibility::Public,
                        "private" => mods.visibility = Visibility::Private,
                        "protected" => mods.visibility = Visibility::Protected,
                        "internal" => mods.visibility = Visibility::Internal,
                        s if s.starts_with("static") => mods.is_static = true,
                        s if s.starts_with("abstract") => mods.is_abstract = true,
                        s if s.starts_with("virtual") => mods.is_virtual = true,
                        s if s.starts_with("override") => mods.is_override = true,
                        s if s.starts_with("readonly") => mods.is_readonly = true,
                        s if s.starts_with("async") => {} // handled in method
                        _ => {}
                    }
                }
            }
            _ => member_pair = Some(p),
        }
    }

    let mp = member_pair.ok_or("Empty class member")?;
    match mp.as_rule() {
        Rule::constructor_declaration => walk_constructor(mp, mods),
        Rule::property_declaration => walk_property(mp, mods),
        Rule::event_declaration => walk_event(mp),
        Rule::method_declaration => walk_method(mp, mods),
        Rule::field_declaration => walk_field(mp, mods),
        other => Err(format!("Unexpected class member: {:?}", other)),
    }
}

fn walk_constructor(pair: Pair<Rule>, _mods: Modifiers) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {} // constructor name (same as class)
            Rule::param_list => params = walk_params(p)?,
            Rule::constructor_initializer => {
                let mut args = Vec::new();
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::argument_list {
                        args = walk_arguments(cp)?;
                    }
                }
                base_args = Some(args.into_iter().map(|a| a.value).collect());
            }
            Rule::block_statement => body = walk_body(p)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        params,
        body,
        base_args,
        visibility: Visibility::Public,
    })
}

fn walk_property(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut getter = None;
    let mut setter = None;
    let mut is_auto = true;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip type
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::property_body => {
                for acc in p.into_inner() {
                    if acc.as_rule() == Rule::accessor {
                        let mut is_get = false;
                        let mut acc_body = None;
                        for ap in acc.into_inner() {
                            match ap.as_rule() {
                                Rule::block_statement => {
                                    acc_body = Some(walk_body(ap)?);
                                    is_auto = false;
                                }
                                Rule::class_modifiers => {} // skip accessor modifiers
                                _ => {
                                    match ap.as_str() {
                                        "get" => is_get = true,
                                        "set" => is_get = false,
                                        _ => {}
                                    }
                                }
                            }
                        }
                        if is_get {
                            getter = acc_body.or(Some(Vec::new()));
                        } else {
                            let param = Param {
                                name: "value".into(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            };
                            setter = Some(PropertySetter {
                                param,
                                body: acc_body.unwrap_or_default(),
                            });
                        }
                    }
                }
            }
            _ => {} // skip initializer expression
        }
    }

    Ok(ClassMember::Property {
        name,
        type_hint: None,
        getter,
        setter,
        is_auto,
        modifiers: mods,
    })
}

fn walk_event(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            _ => {}
        }
    }
    Ok(ClassMember::Event {
        name,
        type_hint,
        params: Vec::new(),
        visibility: Visibility::Public,
    })
}

fn walk_method(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut return_type = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut is_async = false;

    // Check modifiers for async
    if mods.is_abstract { /* abstract methods have no body */ }

    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut iter = parts.into_iter();

    // First is return type
    if let Some(p) = iter.next() {
        if p.as_rule() == Rule::type_name {
            let rt = p.as_str().to_string();
            if rt.starts_with("async") {
                is_async = true;
            }
            return_type = Some(rt);
        }
    }

    // Second is name
    if let Some(p) = iter.next() {
        if p.as_rule() == Rule::ident_name {
            name = p.as_str().to_string();
        }
    }

    // Rest: params and body
    for p in iter {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::block_statement => body = walk_body(p)?,
            _ => {}
        }
    }

    let is_sub = return_type.as_deref() == Some("void");

    let is_generator = body_has_yield(&body);
    Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: mods,
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub,
    }))))
}

fn walk_field(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip type
            Rule::var_declarator_list => {
                // Take first declarator
                if let Some(vd) = p.into_inner().find(|p| p.as_rule() == Rule::var_declarator) {
                    let decl = walk_var_declarator(vd)?;
                    if let BindingPattern::Ident(n) = decl.pattern {
                        name = n;
                    }
                    init = decl.init;
                }
            }
            _ => {}
        }
    }

    Ok(ClassMember::Field {
        name,
        type_hint: None,
        init,
        modifiers: mods,
        with_events: false,
        array_bounds: None,
    })
}

// ── Struct ──────────────────────────────────────────────────────────────────

fn walk_struct_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {} // skip modifiers
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::base_list => {
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        interfaces.push(bp.as_str().to_string());
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.push(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::StructDecl {
        name,
        interfaces,
        members,
        visibility: Visibility::Public,
    })
}

// ── Interface ───────────────────────────────────────────────────────────────

fn walk_interface_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::base_list => {
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        parents.push(bp.as_str().to_string());
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::interface_member {
                        if let Ok(member) = walk_interface_member(m) {
                            members.push(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::InterfaceDecl { name, parents, members })
}

fn walk_interface_member(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let mut type_hint = None;
    let mut name = String::new();
    let mut has_params = false;
    let mut params = Vec::new();
    let mut has_getter = false;
    let mut has_setter = false;
    let src = pair.as_str().to_string();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => { has_params = true; params = walk_params(p)?; }
            _ => {
                let s = p.as_str();
                if s == "get" { has_getter = true; }
                if s == "set" { has_setter = true; }
            }
        }
    }

    if has_params || src.contains('(') {
        let is_sub = type_hint.as_deref() == Some("void");
        Ok(InterfaceMember::Method {
            name,
            params,
            return_type: type_hint,
            is_sub,
        })
    } else {
        Ok(InterfaceMember::Property {
            name,
            type_hint,
            is_readonly: has_getter && !has_setter,
            is_writeonly: !has_getter && has_setter,
        })
    }
}

// ── Enum ────────────────────────────────────────────────────────────────────

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::enum_member_list => {
                for em in p.into_inner() {
                    if em.as_rule() == Rule::enum_member {
                        let mut en = String::new();
                        let mut val = None;
                        for ep in em.into_inner() {
                            match ep.as_rule() {
                                Rule::ident_name => en = ep.as_str().to_string(),
                                _ => val = Some(walk_expression(ep)?),
                            }
                        }
                        members.push(EnumMember { name: en, value: val });
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::EnumDecl { name, members, visibility: Visibility::Public })
}

// ── Record ──────────────────────────────────────────────────────────────────

fn walk_record_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::base_list => {
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        parents.push(bp.as_str().to_string());
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.push(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // Record positional params become fields + constructor
    for param in &params {
        members.push(ClassMember::Field {
            name: param.name.clone(),
            type_hint: param.type_hint.clone(),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        });
    }

    // Generate constructor from params
    if !params.is_empty() {
        let ctor_body: Vec<Statement> = params.iter().map(|p| {
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: p.name.clone(),
                    null_safe: false,
                })],
                value: Expression::ident(&p.name),
            })
        }).collect();
        members.push(ClassMember::Constructor {
            params: params.clone(),
            body: ctor_body,
            base_args: None,
            visibility: Visibility::Public,
        });
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
    })
}

// ── Delegate ────────────────────────────────────────────────────────────────

fn walk_delegate_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut return_type = None;
    let mut params = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::DelegateDecl {
        name,
        params,
        return_type: return_type.clone(),
        is_sub: return_type.as_deref() == Some("void"),
        visibility: Visibility::Public,
    })
}

// ── Control flow ────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("if: no cond")?)?;
    let then_body = vec![walk_statement(inner.next().ok_or("if: no body")?)?];
    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::else_if_clause => {
                let mut eip = p.into_inner();
                let ec = walk_expression(eip.next().ok_or("elif: no cond")?)?;
                let eb = vec![walk_statement(eip.next().ok_or("elif: no body")?)?];
                elifs.push((ec, eb));
            }
            Rule::else_clause => {
                let body = p.into_inner().next().ok_or("else: no body")?;
                else_body = Some(vec![walk_statement(body)?]);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If { cond, then_body, elifs, else_body })
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init = None;
    let mut cond = None;
    let mut update = None;
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_init => {
                let inner = p.into_inner().next().ok_or("Empty for init")?;
                match inner.as_rule() {
                    Rule::local_var_declaration_no_semi => {
                        let mut vi = inner.into_inner();
                        let _type = vi.next(); // skip type/var
                        let mut decls = Vec::new();
                        for d in vi {
                            if d.as_rule() == Rule::var_declarator_list {
                                for vd in d.into_inner() {
                                    if vd.as_rule() == Rule::var_declarator {
                                        decls.push(walk_var_declarator(vd)?);
                                    }
                                }
                            }
                        }
                        init = Some(Box::new(Statement::new(StmtKind::VarDecl {
                            declarations: decls,
                            kind: VarDeclKind::Let,
                        })));
                    }
                    Rule::expression_list => {
                        let first_expr = inner.into_inner().next()
                            .ok_or("Empty expr list")?;
                        let expr = walk_expression(first_expr)?;
                        init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                    }
                    _ => {
                        let expr = walk_expression(inner)?;
                        init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                    }
                }
            }
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(p)?);
                }
            }
            Rule::for_update => {
                // for_update may contain expression_list with multiple expressions
                let inner = p.into_inner().next().ok_or("Empty for update")?;
                if inner.as_rule() == Rule::expression_list {
                    let mut exprs: Vec<Pair<Rule>> = inner.into_inner().collect();
                    if exprs.len() == 1 {
                        update = Some(walk_expression(exprs.remove(0))?);
                    } else {
                        // Multiple update expressions → sequence
                        let seq: Vec<Expression> = exprs.into_iter()
                            .map(walk_expression)
                            .collect::<Result<Vec<_>, _>>()?;
                        update = Some(Expression::new(ExprKind::Sequence(seq)));
                    }
                } else {
                    update = Some(walk_expression(inner)?);
                }
            }
            _ => {
                if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    Ok(StmtKind::For { init, cond, update, body })
}

fn walk_foreach(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var = String::new();
    let mut iter = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw | Rule::type_name => {} // skip type
            Rule::ident_name => var = p.as_str().to_string(),
            Rule::in_kw => {} // skip keyword
            _ => {
                if body.is_empty() {
                    if let Ok(expr) = walk_expression(p.clone()) {
                        iter = expr;
                    } else if let Ok(stmt) = walk_statement(p) {
                        body = vec![stmt];
                    }
                } else if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter,
        body,
        of: true, // foreach is like for-of
        else_body: None,
        is_async: false,
    })
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("while: no cond")?)?;
    let body = vec![walk_statement(inner.next().ok_or("while: no body")?)?];
    Ok(StmtKind::While { cond, body, else_body: None })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body = vec![walk_statement(inner.next().ok_or("do: no body")?)?];
    let cond = walk_expression(inner.next().ok_or("do: no cond")?)?;
    Ok(StmtKind::DoWhile { body, cond, until: false })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(inner.next().ok_or("switch: no expr")?)?;
    let mut cases = Vec::new();
    let mut default = None;

    for p in inner {
        if p.as_rule() == Rule::switch_section {
            let mut labels = Vec::new();
            let mut stmts = Vec::new();
            let mut is_default = false;

            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::switch_label => {
                        let label_src = sp.as_str().trim();
                        if label_src.starts_with("default") {
                            is_default = true;
                        } else {
                            // "case expr:"
                            if let Some(expr_pair) = sp.into_inner().next() {
                                labels.push(walk_expression(expr_pair)?);
                            }
                        }
                    }
                    _ => {
                        if let Ok(stmt) = walk_statement(sp) {
                            stmts.push(stmt);
                        }
                    }
                }
            }

            if is_default {
                default = Some(stmts);
            } else {
                let conditions = labels.into_iter().map(CaseCondition::Value).collect();
                cases.push(SwitchCase { conditions, body: stmts });
            }
        }
    }

    Ok(StmtKind::Switch { expr, cases, default })
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .next()
        .map(walk_expression)
        .transpose()?;
    Ok(StmtKind::Return(expr))
}

/// C# `yield return expr;` → `StmtKind::Expr(Yield(expr))`
/// `yield break;`          → `StmtKind::Return(None)` (ends the coroutine)
fn walk_yield_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let s = pair.as_str();
    let inner = pair.into_inner().next();
    if s.trim_start().starts_with("yield") && s.contains("return") {
        let expr = inner.map(walk_expression).transpose()?;
        let yield_expr = Expression::new(ExprKind::Yield(expr.map(Box::new)));
        Ok(StmtKind::Expr(yield_expr))
    } else {
        // yield break → end the generator
        Ok(StmtKind::Return(None))
    }
}

/// Walk a block scanning for any `yield return` / `yield break` —
/// determines whether the enclosing method is a generator.
fn body_has_yield(body: &[Statement]) -> bool {
    fn expr_has_yield(e: &Expression) -> bool {
        matches!(&e.kind, ExprKind::Yield(_) | ExprKind::YieldFrom(_))
    }
    for s in body {
        match &s.kind {
            StmtKind::Expr(e) if expr_has_yield(e) => return true,
            StmtKind::If { then_body, elifs, else_body, .. } => {
                if body_has_yield(then_body) { return true; }
                for (_, b) in elifs { if body_has_yield(b) { return true; } }
                if let Some(b) = else_body { if body_has_yield(b) { return true; } }
            }
            StmtKind::While { body: b, .. } | StmtKind::ForIn { body: b, .. }
            | StmtKind::For { body: b, .. } | StmtKind::DoWhile { body: b, .. } => {
                if body_has_yield(b) { return true; }
            }
            StmtKind::Try { body: b, catches, finally, .. } => {
                if body_has_yield(b) { return true; }
                for c in catches { if body_has_yield(&c.body) { return true; } }
                if let Some(f) = finally { if body_has_yield(f) { return true; } }
            }
            StmtKind::Block(b) => { if body_has_yield(b) { return true; } }
            _ => {}
        }
    }
    false
}

fn walk_throw(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .next()
        .map(walk_expression)
        .transpose()?;
    Ok(StmtKind::Throw { expr, cause: None })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => body = walk_body(p)?,
            Rule::catch_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::type_name => types.push(cp.as_str().to_string()),
                        Rule::ident_name => var_name = Some(cp.as_str().to_string()),
                        Rule::block_statement => catch_body = walk_body(cp)?,
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
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_body(fp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try { body, catches, else_body: None, finally })
}

fn walk_using_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var = String::new();
    let mut resource = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw | Rule::type_name => {}
            Rule::ident_name => var = p.as_str().to_string(),
            _ => {
                if body.is_empty() {
                    if let Ok(expr) = walk_expression(p.clone()) {
                        resource = expr;
                    } else if let Ok(stmt) = walk_statement(p) {
                        body = vec![stmt];
                    }
                } else if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    Ok(StmtKind::Using { var, resource, body })
}

fn walk_lock(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(inner.next().ok_or("lock: no expr")?)?;
    let body = vec![walk_statement(inner.next().ok_or("lock: no body")?)?];
    Ok(StmtKind::Lock { expr, body })
}

// ── Parameters ──────────────────────────────────────────────────────────────

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::param)
        .map(walk_param)
        .collect()
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut name = String::new();
    let mut type_hint = None;
    let mut default = None;
    let mut pass_by = PassBy::Value;
    let mut is_rest = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_modifier => {
                match p.as_str() {
                    "ref" => pass_by = PassBy::Ref,
                    "out" => pass_by = PassBy::Out,
                    "params" => is_rest = true,
                    _ => {}
                }
            }
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            _ => default = Some(walk_expression(p)?),
        }
    }

    Ok(Param {
        name,
        type_hint,
        default,
        pass_by,
        is_rest,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
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
            let s = pair.as_str();
            // Strip numeric suffix
            let s = s.trim_end_matches(|c: char| c.is_ascii_alphabetic());
            if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(s.parse().map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0b") || s.starts_with("0B") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?)))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(unquote(pair.as_str())))),
        Rule::verbatim_string => {
            let s = pair.as_str();
            // @"..." → strip @" and trailing "
            let inner = &s[2..s.len()-1];
            Ok(ExprKind::Lit(Literal::Str(inner.replace("\"\"", "\""))))
        }
        Rule::char_literal => {
            let s = pair.as_str();
            let inner = &s[1..s.len()-1];
            let ch = match inner {
                "\\n" => '\n',
                "\\t" => '\t',
                "\\r" => '\r',
                "\\0" => '\0',
                "\\\\" => '\\',
                "\\'" => '\'',
                _ => inner.chars().next().unwrap_or('\0'),
            };
            Ok(ExprKind::Lit(Literal::Char(ch)))
        }

        // Keywords
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::this_kw => Ok(ExprKind::This),
        Rule::base_kw => Ok(ExprKind::Super),

        Rule::ident_name | Rule::ident_or_keyword | Rule::primitive_type_ident => {
            let name = pair.as_str();
            match name {
                "true" => Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => Ok(ExprKind::Lit(Literal::Null)),
                "this" => Ok(ExprKind::This),
                "base" => Ok(ExprKind::Super),
                _ => Ok(ExprKind::Ident(name.to_string())),
            }
        }

        // Expression wrapper
        Rule::expression => {
            let inner = pair.into_inner().next().ok_or("Empty expression")?;
            walk_expr_kind(inner)
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
                    let op = match op_str {
                        "+=" => CompoundOp::Add, "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul, "/=" => CompoundOp::Div,
                        "%=" => CompoundOp::Mod,
                        "&=" => CompoundOp::BitAnd, "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor, "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr,
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

        // Lambda
        Rule::lambda_expression => {
            let mut params = Vec::new();
            let mut body = LambdaBody::Expr(Box::new(Expression::null()));
            let mut is_async = false;

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::ident_name => params = vec![Param {
                        name: p.as_str().to_string(),
                        type_hint: None, default: None,
                        pass_by: PassBy::Value, is_rest: false,
                        is_kwargs: false, is_optional: false, is_nullable: false,
                    }],
                    Rule::param_list => params = walk_params(p)?,
                    Rule::lambda_body => {
                        let inner = p.into_inner().next().ok_or("Empty lambda body")?;
                        body = match inner.as_rule() {
                            Rule::block_statement => LambdaBody::Block(walk_body(inner)?),
                            _ => LambdaBody::Expr(Box::new(walk_expression(inner)?)),
                        };
                    }
                    _ => {}
                }
            }

            Ok(ExprKind::Lambda { params, body, is_async, captures: Vec::new() })
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
        Rule::null_coalesce_expr => walk_binary_chain(pair),
        Rule::logical_or | Rule::logical_and => walk_binary_chain(pair),
        Rule::bitwise_or | Rule::bitwise_xor | Rule::bitwise_and => walk_binary_chain(pair),
        Rule::equality => walk_binary_chain(pair),
        Rule::relational => walk_relational(pair),
        Rule::additive | Rule::multiplicative => walk_binary_chain(pair),

        // Unary
        Rule::unary => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty unary")?;
            if first.as_rule() == Rule::postfix {
                return walk_expr_kind(first);
            }
            let op_str = first.as_str().trim();
            let operand = walk_expression(inner.next().ok_or("Missing unary operand")?)?;
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
            let has_postfix = inner.iter().any(|p| p.as_rule() == Rule::postfix_op);
            if !has_postfix { return Ok(base.kind); }
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

        // New expression
        Rule::new_expression => walk_new_expr(pair),

        // Primary
        Rule::primary => {
            let inner = pair.into_inner().next().ok_or("Empty primary")?;
            walk_expr_kind(inner)
        }

        // typeof(Type) → push type name as string
        Rule::typeof_expression => {
            let type_name = pair.into_inner()
                .find(|p| p.as_rule() == Rule::type_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(ExprKind::Lit(Literal::Str(type_name)))
        }

        // nameof(member) → push name as string
        Rule::nameof_expression => {
            let name = pair.into_inner()
                .find(|p| p.as_rule() == Rule::dotted_name)
                .map(|p| {
                    let s = p.as_str();
                    // nameof returns the last segment
                    s.rsplit('.').next().unwrap_or(s).to_string()
                })
                .unwrap_or_default();
            Ok(ExprKind::Lit(Literal::Str(name)))
        }

        // default or default(Type)
        Rule::default_expression => {
            // default(int) → 0, default(bool) → false, etc.
            let type_name = pair.into_inner()
                .find(|p| p.as_rule() == Rule::type_name)
                .map(|p| p.as_str().to_string());
            match type_name.as_deref() {
                Some("int") | Some("long") | Some("short") | Some("byte")
                | Some("double") | Some("float") | Some("decimal") => {
                    Ok(ExprKind::Lit(Literal::Int(0)))
                }
                Some("bool") => Ok(ExprKind::Lit(Literal::Bool(false))),
                Some("char") => Ok(ExprKind::Lit(Literal::Char('\0'))),
                _ => Ok(ExprKind::Lit(Literal::Null)),
            }
        }

        // checked/unchecked → just evaluate inner
        Rule::checked_expression => {
            let inner = pair.into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .ok_or("Empty checked expression")?;
            walk_expr_kind(inner)
        }

        // Interpolated string — parsed atomically, split manually
        Rule::interpolated_string => {
            let s = pair.as_str();
            // Strip $" prefix and " suffix
            let inner = &s[2..s.len()-1];
            let parts = parse_interpolated_parts(inner)?;
            Ok(ExprKind::Interpolation(parts))
        }

        // Type name used as expression (cast, etc.)
        Rule::type_name | Rule::base_type | Rule::dotted_name => {
            Ok(ExprKind::Ident(pair.as_str().to_string()))
        }

        // Passthrough wrappers
        Rule::call_chain => {
            let inner = pair.into_inner().next().ok_or("Empty wrapper")?;
            walk_expr_kind(inner)
        }

        // C# tuple literal: (1, "x", true) → canonical Tuple AST node
        Rule::tuple_literal => {
            let elems: Vec<Expression> = pair.into_inner()
                .map(walk_expression)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(ExprKind::Tuple(elems))
        }

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Binary chain walker ─────────────────────────────────────────────────────

/// Walk relational expression: additive ~ (type_test | binary_relational)*
fn walk_relational(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    let mut left = walk_expression(inner.remove(0))?;

    for p in inner {
        match p.as_rule() {
            Rule::type_test => {
                let mut tt_inner = p.into_inner();
                let kw = tt_inner.next().ok_or("type_test: missing keyword")?;
                let tn = tt_inner.next().ok_or("type_test: missing type")?;
                let type_name = tn.as_str().trim().to_string();
                if kw.as_rule() == Rule::is_kw {
                    left = Expression::new(ExprKind::IsType { expr: Box::new(left), type_name });
                } else {
                    left = Expression::new(ExprKind::Cast { expr: Box::new(left), type_name });
                }
            }
            Rule::binary_relational => {
                let mut br_inner: Vec<Pair<Rule>> = p.into_inner().collect();
                if br_inner.len() >= 2 {
                    let op_str = br_inner[0].as_str().trim();
                    let right = walk_expression(br_inner.remove(1))?;
                    let bin_op = match op_str {
                        "<=" => BinOp::LtEq, ">=" => BinOp::GtEq,
                        "<" => BinOp::Lt, ">" => BinOp::Gt,
                        ">>" => BinOp::Shr, "<<" => BinOp::Shl,
                        _ => BinOp::Lt,
                    };
                    left = Expression::new(ExprKind::Binary {
                        op: bin_op, left: Box::new(left), right: Box::new(right),
                    });
                }
            }
            _ => {
                // Direct operand — shouldn't happen but try as additive
                let right = walk_expression(p)?;
                left = Expression::new(ExprKind::Binary {
                    op: BinOp::Lt, left: Box::new(left), right: Box::new(right),
                });
            }
        }
    }

    Ok(left.kind)
}

fn walk_binary_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    let mut left = walk_expression(inner.remove(0))?;
    let mut i = 0;

    while i + 1 < inner.len() {
        let op_pair = &inner[i];
        let op_str = op_pair.as_str().trim();

        // Check for is/as type operators
        if op_str.starts_with("is ") || op_str.starts_with("is\t") || op_str == "is" {
            // is Type → IsType
            let right = &inner[i + 1];
            let type_name = right.as_str().trim().to_string();
            left = Expression::new(ExprKind::IsType {
                expr: Box::new(left),
                type_name,
            });
            i += 2;
            continue;
        }
        if op_str.starts_with("as ") || op_str.starts_with("as\t") || op_str == "as" {
            let right = &inner[i + 1];
            let type_name = right.as_str().trim().to_string();
            left = Expression::new(ExprKind::Cast {
                expr: Box::new(left),
                type_name,
            });
            i += 2;
            continue;
        }

        let right = walk_expression(inner[i + 1].clone())?;

        let bin_op = match op_str {
            "??" => BinOp::NullCoalesce,
            "||" => BinOp::Or, "&&" => BinOp::And,
            "|" => BinOp::BitOr, "^" => BinOp::BitXor, "&" => BinOp::BitAnd,
            "==" => BinOp::Eq, "!=" => BinOp::NotEq,
            "<" => BinOp::Lt, ">" => BinOp::Gt,
            "<=" => BinOp::LtEq, ">=" => BinOp::GtEq,
            ">>" => BinOp::Shr, "<<" => BinOp::Shl,
            "+" => BinOp::Add, "-" => BinOp::Sub,
            "*" => BinOp::Mul, "/" => BinOp::Div, "%" => BinOp::Mod,
            _ => {
                // relational_op may contain "is Type" or "as Type" as combined text
                if op_str.starts_with("is") {
                    let type_name = op_str.trim_start_matches("is").trim().to_string();
                    left = Expression::new(ExprKind::IsType { expr: Box::new(left), type_name });
                    i += 2;
                    continue;
                }
                if op_str.starts_with("as") {
                    let type_name = op_str.trim_start_matches("as").trim().to_string();
                    left = Expression::new(ExprKind::Cast { expr: Box::new(left), type_name });
                    i += 2;
                    continue;
                }
                BinOp::Add
            }
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
        let chain_src = chain.as_str();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with("?.") {
            // Null-conditional member access
            let name = chain_inner.into_iter()
                .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr), field: name, null_safe: true,
            });
        } else if chain_src.starts_with("(") {
            // Call — normalize known method calls to canonical builtins
            let args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                walk_arguments(arg_pair)?
            } else { Vec::new() };
            expr = canonicalize_method_call(expr, args);
        } else if chain_src.starts_with(".") {
            // Member access — normalize known property accessors to canonical builtins
            let name = chain_inner.into_iter()
                .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            // Canonicalize C# property accessors: Length, Count → __len__
            expr = canonicalize_member_access(expr, &name);
        } else if chain_src.starts_with("[") {
            // Index or range slice. The grammar is `[ expression (".." expression?)? ]`.
            // If `..` is present in the source, build an Index over a Range so the
            // compiler emits a slice via array_slice (standard WASM opcode).
            let exprs: Vec<Pair<Rule>> = chain_inner.into_iter().collect();
            let has_range = chain_src.contains("..");
            if has_range {
                let mut iter = exprs.into_iter();
                let start = iter.next()
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or_else(Expression::null);
                let end = iter.next()
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or_else(|| Expression::int(i32::MAX as i64));
                let range = Expression::new(ExprKind::Range {
                    start: Box::new(start),
                    end: Box::new(end),
                    inclusive: false,
                });
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(range),
                    null_safe: false,
                });
            } else if let Some(idx_pair) = exprs.into_iter().next() {
                let index = walk_expression(idx_pair)?;
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr), index: Box::new(index), null_safe: false,
                });
            }
        }
    }

    Ok(expr.kind)
}

// ── New expression walker ───────────────────────────────────────────────────

fn walk_new_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut type_name = String::new();
    let mut args = Vec::new();
    let mut is_array = false;
    let mut array_init = Vec::new();
    let mut obj_init: Vec<(String, Expression)> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name_for_new => {
                // Strip generic params: "List<string>" → "List"
                let raw = p.as_str();
                type_name = raw.split('<').next().unwrap_or(raw).trim().to_string();
            }
            Rule::argument_list => args = walk_arguments(p)?,
            Rule::array_initializer => {
                is_array = true;
                for ap in p.into_inner() {
                    if let Ok(expr) = walk_expression(ap) {
                        array_init.push(expr);
                    }
                }
            }
            Rule::object_initializer => {
                for ip in p.into_inner() {
                    if ip.as_rule() == Rule::initializer_member {
                        let mut name = String::new();
                        let mut val = Expression::null();
                        for mp in ip.into_inner() {
                            match mp.as_rule() {
                                Rule::ident_name => name = mp.as_str().to_string(),
                                _ => val = walk_expression(mp).unwrap_or(Expression::null()),
                            }
                        }
                        obj_init.push((name, val));
                    }
                }
            }
            _ => {
                // Expression inside brackets for array size
                if let Ok(_expr) = walk_expression(p) {
                    if !is_array {
                        is_array = true;
                    }
                    // Array size expression — ignore for now (dynamic arrays)
                }
            }
        }
    }

    if is_array && !array_init.is_empty() {
        // Array initializer: new[] { 1, 2, 3 } or new int[] { 1, 2, 3 }
        let elements = array_init.into_iter()
            .map(|v| ArrayElement { key: None, value: v, spread: false, by_ref: false })
            .collect();
        return Ok(ExprKind::Array(elements));
    }

    // Build class expression — dotted names become Member chains
    // (e.g. "MyApp.Foo" → Member { Ident("MyApp"), "Foo" })
    let class_expr = build_dotted_expr(&type_name);

    if is_array {
        // new int[5] — create empty array
        return Ok(ExprKind::New {
            class: Box::new(class_expr),
            args: args,
        });
    }

    Ok(ExprKind::New {
        class: Box::new(class_expr),
        args,
    })
}

/// Convert a dotted name like "MyApp.Foo.Bar" into a Member chain expression.
fn build_dotted_expr(name: &str) -> Expression {
    let parts: Vec<&str> = name.split('.').collect();
    if parts.len() == 1 {
        return Expression::ident(parts[0]);
    }
    let mut expr = Expression::ident(parts[0]);
    for part in &parts[1..] {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: part.to_string(),
            null_safe: false,
        });
    }
    expr
}

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::argument)
        .map(|p| {
            let src = p.as_str().trim();
            let by_ref = src.starts_with("ref ") || src.starts_with("out ");
            let inner = p.into_inner().next().ok_or("Empty argument".to_string())?;
            let value = walk_expression(inner)?;
            Ok(Argument { value, name: None, by_ref, spread: false })
        })
        .collect()
}

// ── Helpers ─────────────────────────────────────────────────────────────────

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner()
        .map(walk_statement)
        .collect()
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

fn unquote(s: &str) -> String {
    if s.len() < 2 { return s.to_string(); }
    let inner = &s[1..s.len()-1];
    inner.replace("\\\"", "\"").replace("\\\\", "\\")
        .replace("\\n", "\n").replace("\\t", "\t")
        .replace("\\r", "\r").replace("\\0", "\0")
}

/// Canonicalize C# property/member access to unified AST representation.
/// Normalizes language-specific names to canonical builtin calls so the compiler
/// dispatches uniformly across all languages.
///
/// `arr.Length` → `Call(__len__, [arr])`
/// `list.Count` → `Call(__len__, [list])`
fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    let canonical = match name {
        "Length" | "Count" => Some("__len__"),
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

// C# method call canonicalization is intentionally minimal:
// Methods like .ToString() may be overridden on user classes, so we leave them as
// regular method calls and let the compiler dispatch via the class method binding.
// Only true builtin property accessors like .Length, .Count (handled in
// canonicalize_member_access) are normalized to canonical builtins.
//
// Static helper rewrites are different — `string.Join(sep, arr)` is C# surface
// syntax for what other languages express as `arr.join(sep)`. We rewrite it to
// the canonical instance-method form so the compiler dispatches it through the
// shared value-method path with the correct `this` arg ordering.
fn canonicalize_method_call(callee: Expression, args: Vec<Argument>) -> Expression {
    // Static method rewrites to canonical instance form
    if let ExprKind::Member { object, field, .. } = &callee.kind {
        if let ExprKind::Ident(obj_name) = &object.kind {
            // string.Join(sep, arr) → arr.join(sep)
            if obj_name.eq_ignore_ascii_case("string")
               && field.eq_ignore_ascii_case("Join")
               && args.len() == 2
            {
                let sep = args[0].value.clone();
                let arr = args[1].value.clone();
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(arr),
                        field: "join".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(sep)],
                    optional: false,
                });
            }
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

/// Parse interpolated string parts from the raw content between $" and "
fn parse_interpolated_parts(s: &str) -> Result<Vec<InterpolPart>, String> {
    let mut parts = Vec::new();
    let mut text = String::new();
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        if chars[i] == '{' {
            // Flush text
            if !text.is_empty() {
                parts.push(InterpolPart::Text(text.clone()));
                text.clear();
            }
            // Find matching }
            let start = i + 1;
            let mut depth = 1;
            i += 1;
            while i < chars.len() && depth > 0 {
                if chars[i] == '{' { depth += 1; }
                if chars[i] == '}' { depth -= 1; }
                if depth > 0 { i += 1; }
            }
            let expr_str: String = chars[start..i].iter().collect();
            // Parse the expression
            let module = super::CSharpParser::parse(super::Rule::expression, &expr_str)
                .map_err(|e| format!("Interpolation parse error: {}", e))?;
            let expr_pair = module.into_iter().next().ok_or("Empty interpolation expression")?;
            let expr = walk_expression(expr_pair)?;
            parts.push(InterpolPart::Expr(expr));
            i += 1; // skip }
        } else {
            text.push(chars[i]);
            i += 1;
        }
    }

    if !text.is_empty() {
        parts.push(InterpolPart::Text(text));
    }

    Ok(parts)
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
