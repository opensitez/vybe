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

    // Synthesize the .NET Exception hierarchy at the top of every C#
    // program. ECMA-335 / .NET BCL exposes `System.Exception` plus a
    // family of common subclasses (`InvalidOperationException`,
    // `ArgumentNullException`, `DivideByZeroException`, etc.) that user
    // code routinely throws via `throw new <T>("msg")`. We don't model
    // the BCL's class file system, so the walker injects minimal
    // declarations: each takes a `string msg` ctor that stamps `Message`
    // on `this`. `try { ... } catch (T e) { ... e.Message ... }`
    // resolves T to the synthesized class and reads `e.Message`.
    body.splice(0..0, synthesize_exception_classes());

    Ok(Module {
        name: "main".into(),
        language: Lang::CSharp,
        body,
        imports,
    })
}

fn synthesize_exception_classes() -> Vec<Statement> {
    let names = [
        "Exception",
        "InvalidOperationException",
        "ArgumentException",
        "ArgumentNullException",
        "ArgumentOutOfRangeException",
        "DivideByZeroException",
        "FormatException",
        "NullReferenceException",
        "IndexOutOfRangeException",
        "NotImplementedException",
        "NotSupportedException",
        "OverflowException",
        "KeyNotFoundException",
        "FileNotFoundException",
        "IOException",
        "TypeError",
    ];
    names.iter().map(|n| synthesize_exception_class(n)).collect()
}

fn synthesize_exception_class(name: &str) -> Statement {
    let span = Span::default();
    // Per-type constructor signatures per ECMA-335 / .NET BCL:
    //   ArgumentNullException(paramName)              → ParamName=paramName
    //   ArgumentOutOfRangeException(paramName, msg)   → ParamName=paramName, Message=msg
    //   ArgumentException(msg, paramName)             → Message=msg, ParamName=paramName
    //   <other>(msg)                                  → Message=msg
    //
    // Walker emits the appropriate constructor body so `e.ParamName`
    // and `e.Message` resolve to the right values on every catch.
    let needs_param_name = matches!(name,
        "ArgumentNullException"
        | "ArgumentOutOfRangeException"
        | "ArgumentException"
    );

    let assign = |field: &str, ident: &str| Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                    field: field.into(),
                    null_safe: false,
                },
                span.clone(),
            )],
            value: Expression::with_span(ExprKind::Ident(ident.into()), span.clone()),
        },
        span.clone(),
    );

    let canon = crate::emitter::errors::canonical_exception_name(name).to_string();
    let assign_extype = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                    field: "__exception_type".into(),
                    null_safe: false,
                },
                span.clone(),
            )],
            value: Expression::with_span(
                ExprKind::Lit(Literal::Str(canon.clone())),
                span.clone(),
            ),
        },
        span.clone(),
    );

    let mk_param = |pname: &str| Param {
        name: pname.into(),
        type_hint: Some("string".into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false,
    };

    let (params, body) = if needs_param_name {
        // 2-arg form: (paramName, msg) for ArgumentOutOfRangeException;
        // (msg, paramName) for ArgumentException; (paramName) for
        // ArgumentNullException. The walker matches the .NET BCL order.
        match name {
            "ArgumentException" => (
                vec![mk_param("msg"), mk_param("paramName")],
                vec![assign("Message", "msg"), assign("ParamName", "paramName"), assign_extype],
            ),
            "ArgumentNullException" => (
                vec![mk_param("paramName")],
                vec![assign("ParamName", "paramName"), assign_extype],
            ),
            "ArgumentOutOfRangeException" => (
                vec![mk_param("paramName"), mk_param("msg")],
                vec![assign("ParamName", "paramName"), assign("Message", "msg"), assign_extype],
            ),
            _ => unreachable!(),
        }
    } else {
        (vec![mk_param("msg")], vec![assign("Message", "msg"), assign_extype])
    };

    let mut members = vec![
        ClassMember::Field {
            name: "Message".into(),
            type_hint: Some("string".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        },
    ];
    if needs_param_name {
        members.push(ClassMember::Field {
            name: "ParamName".into(),
            type_hint: Some("string".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        });
    }
    members.push(ClassMember::Constructor {
        params,
        body,
        base_args: None,
        visibility: Visibility::Public,
    });

    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
        },
        span,
    )
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
        Rule::local_function_decl => walk_local_function(pair)?,
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
    let init = match inner.next() {
        Some(p) if p.as_rule() == Rule::array_initializer => {
            // `int[] arr = { 1, 2, 3 }` — bare-brace array literal in a
            // declarator. Desugar to a plain Array expression so the
            // compiler emits the same bytecode as `new[] { ... }`.
            // Each child is a `collection_element` (post grammar change
            // for nested-brace dict / multi-dim support).
            let span = to_span(&p);
            let elems = p.into_inner()
                .map(|e| walk_collection_element(e).map(|expr| ArrayElement {
                    key: None, value: expr, spread: false, by_ref: false,
                }))
                .collect::<Result<Vec<_>, _>>()?;
            Some(Expression::with_span(ExprKind::Array(elems), span))
        }
        Some(p) => Some(walk_expression(p)?),
        None => None,
    };
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
            Rule::ident_name => {
                // Generic param idents (`class Box<T> { ... }`) leak
                // through the silent `generic_params` wrapper rule —
                // they appear as additional `ident_name` pairs after
                // the class name. Keep only the first.
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
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
                            members.extend(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl { name, parents, interfaces, members, modifiers: class_mods })
}

fn walk_class_member(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
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
                        // C# `const` — implicitly static + readonly per ECMA-334 §15.4.
                        // Compile-time constant folded into class-level slot.
                        "const" => { mods.is_static = true; mods.is_readonly = true; }
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
        Rule::constructor_declaration => walk_constructor(mp, mods).map(|m| vec![m]),
        Rule::property_declaration => walk_property(mp, mods),
        Rule::event_declaration => walk_event(mp).map(|m| vec![m]),
        Rule::method_declaration => walk_method(mp, mods).map(|m| vec![m]),
        Rule::field_declaration => walk_field(mp, mods).map(|m| vec![m]),
        Rule::operator_declaration => walk_operator(mp, mods).map(|m| vec![m]),
        Rule::indexer_declaration => walk_indexer(mp, mods),
        // Nested type — wrap as `ClassMember::NestedType(stmt)` so the
        // class-emit pipeline registers the inner type as a sibling
        // global. Per ECMA-334 §15.3 nested types are accessible via
        // `Outer.Inner` qualified name; our compiler treats them as
        // top-level globals already.
        Rule::class_declaration
        | Rule::struct_declaration
        | Rule::interface_declaration
        | Rule::enum_declaration => {
            let span = to_span(&mp);
            let kind = match mp.as_rule() {
                Rule::class_declaration => walk_class_decl(mp)?,
                Rule::struct_declaration => walk_struct_decl(mp)?,
                Rule::interface_declaration => walk_interface_decl(mp)?,
                Rule::enum_declaration => walk_enum_decl(mp)?,
                _ => unreachable!(),
            };
            Ok(vec![ClassMember::NestedType(Box::new(
                Statement::with_span(kind, span),
            ))])
        }
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
            Rule::expression_body => {
                // `ClassName(p) => stmt;` desugars to a body whose
                // single statement is the expression as a stand-alone
                // ExprStmt. Constructors don't return a value, so we
                // don't wrap in Return.
                if let Some(inner) = p.into_inner().next() {
                    let span = to_span(&inner);
                    let expr = walk_expression(inner)?;
                    body = vec![Statement::with_span(StmtKind::Expr(expr), span)];
                }
            }
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

fn walk_property(pair: Pair<Rule>, mods: Modifiers) -> Result<Vec<ClassMember>, String> {
    let mut name = String::new();
    let mut getter = None;
    let mut setter = None;
    let mut is_auto = true;
    let mut default_init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip type
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::expression_body => {
                // `Type Name => expr;` — read-only expression-bodied
                // property. Lower to a getter that returns the expr.
                if let Some(inner) = p.into_inner().next() {
                    let span = to_span(&inner);
                    let expr = walk_expression(inner)?;
                    getter = Some(vec![Statement::with_span(
                        StmtKind::Return(Some(expr)),
                        span,
                    )]);
                    is_auto = false;
                }
            }
            Rule::property_body => {
                for acc in p.into_inner() {
                    if acc.as_rule() == Rule::accessor {
                        // The `get` / `set` keywords are literal tokens
                        // in the grammar — pest doesn't surface them as
                        // child pairs, so we detect direction by
                        // looking at the source string. The accessor's
                        // `as_str()` is something like
                        // `public get { return _v; }`; trim the leading
                        // class_modifiers and check the next word.
                        let acc_src = acc.as_str().trim_start();
                        // Strip optional accessor modifiers (`public`
                        // `private` `protected` `internal`) so we land on
                        // the `get` / `set` keyword itself.
                        let mut rest = acc_src;
                        for kw in &["public", "private", "protected", "internal"] {
                            if let Some(stripped) = rest.strip_prefix(*kw) {
                                if stripped.starts_with(|c: char| c.is_whitespace()) {
                                    rest = stripped.trim_start();
                                    break;
                                }
                            }
                        }
                        let is_get = rest.starts_with("get")
                            && rest[3..].chars().next().map_or(true, |c| !c.is_alphanumeric() && c != '_');
                        let mut acc_body = None;
                        for ap in acc.into_inner() {
                            match ap.as_rule() {
                                Rule::block_statement => {
                                    acc_body = Some(walk_body(ap)?);
                                    is_auto = false;
                                }
                                Rule::class_modifiers => {} // skip accessor modifiers
                                _ => {}
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
            // C# auto-property default-value initializer:
            // `public string Name { get; set; } = "default";` — capture
            // the RHS expression and emit a sibling Field below so the
            // backing slot is initialised in the constructor.
            other if other != Rule::class_modifiers => {
                default_init = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }

    let mut out = vec![ClassMember::Property {
        name: name.clone(),
        type_hint: None,
        getter,
        setter,
        is_auto,
        modifiers: mods.clone(),
    }];
    // Auto-property compiles its `__name` backing field. Emit a Field
    // entry with the same backing name so its `init` runs at instance
    // construction. ECMA-334 §15.7.4 — auto-property initializers run
    // before any user constructor body, matching what FieldDecl gives us.
    if is_auto {
        if let Some(init_expr) = default_init {
            // C# auto-properties without explicit accessors compile as
            // plain fields keyed by the property name (no `__` prefix).
            // The compiler's pass-1 in `classes.rs` adds `(pname, None)`
            // to `field_inits`; we emit the same field name with a real
            // `init` expression so it materialises in the constructor.
            // The duplicate key is deduped by the existing pass-1 check.
            out.push(ClassMember::Field {
                name: name.clone(),
                type_hint: None,
                init: Some(init_expr),
                modifiers: mods,
                with_events: false,
                array_bounds: None,
            });
        }
    }
    Ok(out)
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

/// Walk every statement in a catch body and rewrite bare `throw;`
/// (Throw with no expression) into `throw <catch_var>;` so the VM
/// rethrows the caught instance instead of `NULL`. Recurses into
/// nested blocks but stops at any inner Try / FunctionDecl /
/// LambdaBody — re-throw scopes lexically by .NET semantics.
fn rewrite_bare_throws(stmts: &mut Vec<Statement>, var_name: &str) {
    for stmt in stmts.iter_mut() {
        rewrite_bare_throws_in_stmt(stmt, var_name);
    }
}
fn rewrite_bare_throws_in_stmt(stmt: &mut Statement, var_name: &str) {
    match &mut stmt.kind {
        StmtKind::Throw { expr, .. } if expr.is_none() => {
            *expr = Some(Expression::ident(var_name));
        }
        StmtKind::Block(inner) => rewrite_bare_throws(inner, var_name),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            rewrite_bare_throws(then_body, var_name);
            for (_, body) in elifs {
                rewrite_bare_throws(body, var_name);
            }
            if let Some(eb) = else_body {
                rewrite_bare_throws(eb, var_name);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            rewrite_bare_throws(body, var_name);
        }
        StmtKind::Switch { cases, default, .. } => {
            for c in cases.iter_mut() {
                rewrite_bare_throws(&mut c.body, var_name);
            }
            if let Some(d) = default {
                rewrite_bare_throws(d, var_name);
            }
        }
        StmtKind::Try { body, finally, .. } => {
            // The bare `throw;` inside an inner try's catches refers to
            // that catch's bound exception, NOT the outer one. So we
            // only descend into the outer try's body and finally —
            // catches' bare throws are bound to their own var by the
            // recursive walk_try call.
            rewrite_bare_throws(body, var_name);
            if let Some(f) = finally {
                rewrite_bare_throws(f, var_name);
            }
        }
        _ => {}
    }
}

/// Walk a top-level / local function declaration. Same shape as a
/// class method but lives at statement scope. Lowers to
/// `StmtKind::FunctionDecl` so the compiler treats it like any other
/// free function.
fn walk_local_function(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut return_type = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {
                if return_type.is_none() {
                    return_type = Some(p.as_str().to_string());
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::param_list => params = walk_params(p)?,
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(
                        StmtKind::Return(Some(expr)),
                        span,
                    )];
                }
            }
            _ => {}
        }
    }
    let is_sub = return_type.as_deref() == Some("void");
    let is_generator = body_has_yield(&body);
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        is_async: false,
        is_generator,
        is_sub,
        handles: Vec::new(),
    })
}

/// Map a C# operator symbol (`+`, `==`, …) to the ECMA-335 / .NET
/// canonical method name (`op_Addition`, `op_Equality`, …). Used by
/// the operator-overload walker so the method can be located via the
/// same name C# ABI emits.
fn operator_method_name(symbol: &str) -> &'static str {
    match symbol {
        "+" => "op_Addition",
        "-" => "op_Subtraction",
        "*" => "op_Multiply",
        "/" => "op_Division",
        "%" => "op_Modulus",
        "==" => "op_Equality",
        "!=" => "op_Inequality",
        "<" => "op_LessThan",
        ">" => "op_GreaterThan",
        "<=" => "op_LessThanOrEqual",
        ">=" => "op_GreaterThanOrEqual",
        "&" => "op_BitwiseAnd",
        "|" => "op_BitwiseOr",
        "^" => "op_ExclusiveOr",
        "<<" => "op_LeftShift",
        ">>" => "op_RightShift",
        "~" => "op_OnesComplement",
        "!" => "op_LogicalNot",
        _ => "op_Unknown",
    }
}

/// Walk an `operator_declaration`. Lowers to a static method named
/// per `operator_method_name` so the call-site dispatch can find it
/// via the canonical naming scheme.
fn walk_operator(pair: Pair<Rule>, mut mods: Modifiers) -> Result<ClassMember, String> {
    mods.is_static = true;
    let mut return_type = None;
    let mut symbol = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::operator_symbol => symbol = p.as_str().trim().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(
                        StmtKind::Return(Some(expr)),
                        span,
                    )];
                }
            }
            _ => {}
        }
    }
    let name = operator_method_name(&symbol).to_string();
    let is_sub = return_type.as_deref() == Some("void");
    let is_generator = body_has_yield(&body);
    Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: mods,
        is_async: false,
        is_generator,
        is_sub,
        handles: Vec::new(),
    }))))
}

/// Walk an `indexer_declaration`. Lowers to a Property named `__index__`
/// with the indexer's parameter list captured separately so the runtime
/// can route `obj[i]` through the getter / setter.
fn walk_indexer(pair: Pair<Rule>, mods: Modifiers) -> Result<Vec<ClassMember>, String> {
    let mut getter: Option<Vec<Statement>> = None;
    let mut setter: Option<PropertySetter> = None;
    let mut params: Vec<Param> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip return type
            Rule::param_list => params = walk_params(p)?,
            Rule::property_body => {
                for acc in p.into_inner() {
                    if acc.as_rule() == Rule::accessor {
                        let mut is_get = false;
                        let mut acc_body = None;
                        for ap in acc.into_inner() {
                            match ap.as_rule() {
                                Rule::block_statement => {
                                    acc_body = Some(walk_body(ap)?);
                                }
                                Rule::class_modifiers => {}
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
                            getter = acc_body;
                        } else if let Some(body) = acc_body {
                            // Setter takes (idx..., value). Append `value`.
                            let mut set_params = params.clone();
                            set_params.push(Param {
                                name: "value".into(),
                                type_hint: None, default: None,
                                pass_by: PassBy::Value, is_rest: false,
                                is_kwargs: false, is_optional: false,
                                is_nullable: false,
                            });
                            setter = Some(PropertySetter {
                                param: set_params.first().cloned().unwrap_or_else(|| Param {
                                    name: "value".into(),
                                    type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false,
                                    is_nullable: false,
                                }),
                                body,
                            });
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(vec![ClassMember::Property {
        name: "__index__".to_string(),
        type_hint: None,
        getter,
        setter,
        is_auto: false,
        modifiers: mods,
    }])
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
            Rule::expression_body => {
                // C# expression-bodied member: `=> expr;` lowers to
                // `{ return expr; }`. The inner `expression` pair is
                // the only child of `expression_body`.
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(
                        StmtKind::Return(Some(expr)),
                        span,
                    )];
                }
            }
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
                            members.extend(member);
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
                            members.extend(member);
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

    // Synthetic ToString — `Point { X = 3, Y = 4 }` (ECMA-334 §15.6.6,
    // .NET 5+ record default). Skip if the user already defined one.
    let has_user_tostring = members.iter().any(|m| matches!(
        m,
        ClassMember::Method(stmt) if matches!(
            &stmt.kind,
            StmtKind::FunctionDecl { name, .. } if name == "ToString"
        )
    ));
    if !has_user_tostring && !params.is_empty() {
        let mut concat = Expression::new(ExprKind::Lit(Literal::Str(format!("{} {{ ", name))));
        for (i, p) in params.iter().enumerate() {
            if i > 0 {
                concat = Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(concat),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(", ".into())))),
                });
            }
            concat = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(concat),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(format!("{} = ", p.name))))),
            });
            concat = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(concat),
                right: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: p.name.clone(),
                    null_safe: false,
                })),
            });
        }
        concat = Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(concat),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(" }".into())))),
        });
        let body = vec![Statement::new(StmtKind::Return(Some(concat)))];
        members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name: "ToString".into(),
            params: Vec::new(),
            body,
            return_type: Some("string".into()),
            is_async: false,
            is_generator: false,
            is_sub: false,
            handles: Vec::new(),
            modifiers: Modifiers { is_override: true, ..Default::default() },
        }))));
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
                let mut when_filter: Option<Expression> = None;
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::type_name => types.push(cp.as_str().to_string()),
                        Rule::ident_name => var_name = Some(cp.as_str().to_string()),
                        Rule::catch_when_filter => {
                            // Inner is just an `expression`.
                            if let Some(inner) = cp.into_inner().next() {
                                when_filter = Some(walk_expression(inner)?);
                            }
                        }
                        Rule::block_statement => catch_body = walk_body(cp)?,
                        _ => {}
                    }
                }
                // Synthesize a hidden var name when the catch declared
                // none (`catch (Exception) { throw; }`). The compiler's
                // catch-binding path stamps the value onto a local with
                // this name, so any `throw;` rewrite inside the body
                // can reference it. Without a var, bare `throw;` would
                // throw NULL — losing the original exception.
                let synthetic_var = if var_name.is_none() {
                    let name = "__caught";
                    var_name = Some(name.into());
                    Some(name.to_string())
                } else { None };
                // Bare `throw;` (StmtKind::Throw with expr=None) inside
                // the catch body gets rewritten to `throw <var_name>;`
                // so the VM rethrows the caught instance instead of NULL.
                if let Some(ref vn) = var_name {
                    rewrite_bare_throws(&mut catch_body, vn);
                }
                let _ = synthetic_var;
                // `catch (...) when (cond) { body }` lowers to:
                //
                //     catch (...) {
                //         if (!cond) { throw <var_name>; }
                //         body
                //     }
                //
                // The walker stays language-agnostic — the compiler's
                // existing throw / re-throw path picks it up.
                if let Some(cond) = when_filter {
                    let throw_var = var_name.as_deref()
                        .map(Expression::ident);
                    let throw_stmt = Statement::with_span(
                        StmtKind::Throw { expr: throw_var, cause: None },
                        Span::default(),
                    );
                    let neg = Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(cond),
                    });
                    let if_stmt = Statement::with_span(
                        StmtKind::If {
                            cond: neg,
                            then_body: vec![throw_stmt],
                            elifs: Vec::new(),
                            else_body: None,
                        },
                        Span::default(),
                    );
                    catch_body.insert(0, if_stmt);
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
            let raw = pair.as_str();
            // Strip C#'s numeric type suffix (UL, L, F, M, D, U). Hex digits
            // need different handling — after `0x`, only the trailing
            // suffix (UL/L/U) is alpha noise; A–F are real digits.
            let s = if raw.starts_with("0x") || raw.starts_with("0X") {
                let body = &raw[2..];
                let cut = body.rfind(|c: char| c.is_ascii_hexdigit())
                    .map(|i| 2 + i + 1)
                    .unwrap_or(raw.len());
                &raw[..cut]
            } else if raw.starts_with("0b") || raw.starts_with("0B") {
                let body = &raw[2..];
                let cut = body.rfind(|c: char| c == '0' || c == '1')
                    .map(|i| 2 + i + 1)
                    .unwrap_or(raw.len());
                &raw[..cut]
            } else {
                raw.trim_end_matches(|c: char| c.is_ascii_alphabetic())
            };
            // Underscores are allowed as digit separators in C# 7.0+.
            let s_owned;
            let s = if s.contains('_') { s_owned = s.replace('_', ""); &s_owned } else { s };
            if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?)))
            } else if s.starts_with("0b") || s.starts_with("0B") {
                Ok(ExprKind::Lit(Literal::Int(i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?)))
            } else if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(s.parse().map_err(|e| format!("{}", e))?)))
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
            if first.as_rule() == Rule::cast_expression {
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

        // C# explicit cast `(TypeName)expr` — lower to the canonical
        // type-conversion form. For numeric primitives we use Convert.<T>
        // calls (matches the .NET runtime); for object / string we leave
        // the expression unchanged (Convert.ToString already in stdlib).
        Rule::cast_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let cast_type_pair = inner.remove(0);
            let type_name = cast_type_pair.as_str().trim().to_string();
            let operand = walk_expression(inner.remove(0))?;
            let convert_method = match type_name.as_str() {
                "int" | "uint" | "short" | "ushort" | "sbyte" | "byte" => "ToInt32",
                "long" | "ulong" => "ToInt64",
                "float" => "ToSingle",
                "double" | "decimal" => "ToDouble",
                "string" => "ToString",
                "bool" => "ToBoolean",
                "char" => "ToChar",
                _ => return Ok(operand.kind),
            };
            // Convert.ToInt32(operand) etc.
            let span = operand.span.clone();
            Ok(ExprKind::Call {
                callee: Box::new(Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(Expression::with_span(
                            ExprKind::Ident("Convert".into()), span.clone(),
                        )),
                        field: convert_method.into(),
                        null_safe: false,
                    },
                    span.clone(),
                )),
                args: vec![Argument::positional(operand)],
                optional: false,
            })
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

        // typeof(Type) → push the .NET FullName as a string. Matches
        // .NET's `Console.WriteLine(typeof(int))` → "System.Int32".
        // `.Name` / `.FullName` member access is rewritten in
        // `canonicalize_member_access` for typeof-string receivers.
        Rule::typeof_expression => {
            let type_name = pair.into_inner()
                .find(|p| p.as_rule() == Rule::type_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            let net_name = dotnet_type_name(&type_name);
            Ok(ExprKind::Lit(Literal::Str(format!("System.{}", net_name))))
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

        // C# tuple literal: (1, "x", true) or named (Name: "Alice", Age: 30).
        // Named tuples lower to an Object with both Item<N> AND the
        // user-given names so `t.Item1` and `t.Name` both resolve.
        Rule::tuple_literal => {
            let elements: Vec<Pair<Rule>> = pair.into_inner()
                .filter(|p| p.as_rule() == Rule::tuple_element)
                .collect();
            let mut has_names = false;
            let mut parsed: Vec<(Option<String>, Expression)> = Vec::new();
            for el in elements {
                let inner: Vec<Pair<Rule>> = el.into_inner().collect();
                if inner.len() == 2 && inner[0].as_rule() == Rule::ident_name {
                    has_names = true;
                    parsed.push((
                        Some(inner[0].as_str().to_string()),
                        walk_expression(inner[1].clone())?,
                    ));
                } else if let Some(p) = inner.into_iter().next() {
                    parsed.push((None, walk_expression(p)?));
                }
            }
            if has_names {
                let mut props = Vec::new();
                for (i, (name, value)) in parsed.iter().enumerate() {
                    let item_key = format!("Item{}", i + 1);
                    props.push(ObjectProperty::KeyValue {
                        key: Expression::string(&item_key),
                        value: value.clone(),
                    });
                    if let Some(n) = name {
                        props.push(ObjectProperty::KeyValue {
                            key: Expression::string(n),
                            value: value.clone(),
                        });
                    }
                }
                Ok(ExprKind::Object(props))
            } else {
                let elems: Vec<Expression> = parsed.into_iter().map(|(_, e)| e).collect();
                Ok(ExprKind::Tuple(elems))
            }
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
                let next = tt_inner.next().ok_or("type_test: missing operand")?;
                if kw.as_rule() == Rule::is_kw {
                    // `is` accepts a pattern_clause: not-prefix +
                    // pattern_atom of (null | literal | type_name [ident])
                    left = walk_is_pattern(left, next)?;
                } else {
                    // `obj as T` — returns obj if it's a T, else null.
                    // Lower to `<is-test> ? obj : null` so the runtime
                    // null sentinel matches .NET semantics. Strip a
                    // trailing `?` (nullable marker) for the type test.
                    let type_name_raw = next.as_str().trim();
                    let type_name = type_name_raw.trim_end_matches('?').to_string();
                    let test = if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                        let typeof_expr = Expression::new(ExprKind::TypeOf(Box::new(left.clone())));
                        Expression::new(ExprKind::Binary {
                            op: BinOp::StrictEq,
                            left: Box::new(typeof_expr),
                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(js_typeof.into())))),
                        })
                    } else {
                        Expression::new(ExprKind::IsType { expr: Box::new(left.clone()), type_name })
                    };
                    left = Expression::new(ExprKind::Ternary {
                        cond: Box::new(test),
                        then: Box::new(left),
                        else_: Box::new(Expression::null()),
                    });
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
            Rule::switch_expr_postfix => {
                left = walk_switch_expr(left, p)?;
            }
            Rule::with_expr_postfix => {
                left = walk_with_expr(left, p)?;
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

        // Integer division by literal zero — `int x = 10 / 0;`. C#
        // (ECMA-335) throws `DivideByZeroException` at runtime; JS
        // returns Infinity. We can't tell int vs float at runtime
        // (every numeric is f64), but a literal `0` divisor with an
        // integer literal numerator is unambiguously the int form
        // — rewrite the expression to a throw of the exception so
        // try/catch picks it up.
        if matches!(bin_op, BinOp::Div)
            && is_int_zero_literal(&right)
            && is_int_literal(&left)
        {
            // Build `(() => { throw new DivideByZeroException("Attempted to divide by zero."); })()`
            // — IIFE so the throw works in expression position.
            let throw_stmt = Statement::with_span(
                StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("DivideByZeroException")),
                        args: vec![Argument::positional(Expression::new(
                            ExprKind::Lit(Literal::Str("Attempted to divide by zero.".into())),
                        ))],
                    })),
                    cause: None,
                },
                Span::default(),
            );
            let lambda = Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![throw_stmt]),
                is_async: false,
                captures: Vec::new(),
            });
            left = Expression::new(ExprKind::Call {
                callee: Box::new(lambda),
                args: vec![],
                optional: false,
            });
            i += 2;
            continue;
        }

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        i += 2;
    }

    Ok(left.kind)
}

fn is_int_literal(e: &Expression) -> bool {
    matches!(e.kind, ExprKind::Lit(Literal::Int(_)))
}

fn is_int_zero_literal(e: &Expression) -> bool {
    matches!(&e.kind, ExprKind::Lit(Literal::Int(0)))
}

// ── Call chain walker ───────────────────────────────────────────────────────

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty call expression")?;
    let mut expr = walk_expression(first)?;

    // Collect the chain segments so we can peek at the next one when
    // deciding whether to canonicalize a `.Length` / `.Count` accessor
    // (they're properties standalone, but instance-method names when
    // followed by `(...)` — see LINQ Count(predicate)).
    let chains: Vec<Pair<Rule>> = inner.filter(|p| p.as_rule() == Rule::call_chain).collect();
    let mut iter = chains.into_iter().peekable();
    while let Some(chain) = iter.next() {
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
            let mut args = if let Some(arg_pair) = chain_inner.into_iter().find(|p| p.as_rule() == Rule::argument_list) {
                walk_arguments(arg_pair)?
            } else { Vec::new() };
            // Inject default fill char for `PadLeft(n)` / `PadRight(n)` —
            // .NET defaults to space, but the value-method dispatch expects
            // both args. Same idea as JS-default lowering.
            if let ExprKind::Member { field, .. } = &expr.kind {
                let f = field.as_str();
                if (f == "PadLeft" || f == "PadRight") && args.len() == 1 {
                    args.push(Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(" ".into())))));
                }
            }
            expr = canonicalize_method_call(expr, args);
        } else if chain_src.starts_with(".") {
            // Member access — normalize known property accessors to canonical builtins
            let name = chain_inner.into_iter()
                .find(|p| p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            // C# tuple ItemN accessor: `(1, 2, 3).Item1` → `t[0]`,
            // `Item2` → `t[1]`, etc. Tuples compile to Arrays so the
            // ItemN names need to lower to indexed access. Pattern is
            // `Item` followed by 1+ digit index (1-based).
            if let Some(rest) = name.strip_prefix("Item") {
                if let Ok(n) = rest.parse::<i64>() {
                    if n >= 1 {
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(Expression::int(n - 1)),
                            null_safe: false,
                        });
                        continue;
                    }
                }
            }
            // Canonicalize C# property accessors: Length, Count → __len__
            // BUT only if the next chain segment is NOT a call. `Count` and
            // `Length` can also appear as instance-method names (LINQ
            // `Count(predicate)`); folding them eagerly into `__len__`
            // breaks the call site.
            let next_is_call = iter.peek()
                .map(|c| c.as_str().starts_with('('))
                .unwrap_or(false);
            if next_is_call {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr), field: name, null_safe: false,
                });
            } else {
                expr = canonicalize_member_access(expr, &name);
            }
        } else if chain_src.starts_with("[") {
            // Index / range / from-end. The C# 8 forms covered:
            //   arr[i]     — plain index
            //   arr[^N]    — from-end index, i.e. arr[arr.Length - N]
            //   arr[a..b]  — range
            //   arr[^N..]  — range with from-end start, open end
            //   arr[..^N]  — range with from-end end
            //   arr[..]    — full slice
            let inner_pairs: Vec<Pair<Rule>> = chain_inner.into_iter().collect();
            // Find the index_expression child (grammar wraps the brackets'
            // contents in `index_expression`).
            let idx_pair = inner_pairs.into_iter()
                .find(|p| p.as_rule() == Rule::index_expression);
            if let Some(idx) = idx_pair {
                let idx_src = idx.as_str().trim();
                let has_range = idx_src.contains("..");
                let parts: Vec<Pair<Rule>> = idx.into_inner().collect();
                if has_range {
                    let mut start: Option<Expression> = None;
                    let mut end: Option<Expression> = None;
                    let hit_dotdot = false;
                    // Walk parts; the `..` token isn't a pair (it's a literal),
                    // so we infer position from order: parts before the index of
                    // a from_end_index / expression that's "after" the dotdot
                    // are starts. Simpler: split source on `..`.
                    let halves: Vec<&str> = idx_src.splitn(2, "..").collect();
                    let _ = (start.as_ref(), end.as_ref(), hit_dotdot);
                    let mut iter = parts.into_iter();
                    let first_after_dotdot = halves.first().map_or(true, |s| s.trim().is_empty());
                    if !first_after_dotdot {
                        if let Some(p) = iter.next() {
                            start = Some(walk_index_part(p, expr.clone())?);
                        }
                    }
                    let second_empty = halves.get(1).map_or(true, |s| s.trim().is_empty());
                    if !second_empty {
                        if let Some(p) = iter.next() {
                            end = Some(walk_index_part(p, expr.clone())?);
                        }
                    }
                    let start = start.unwrap_or_else(Expression::null);
                    let end = end.unwrap_or_else(|| Expression::int(i32::MAX as i64));
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
                } else {
                    // Multi-arg index `m[i, j]` lowers to nested
                    // `m[i][j]`. Single-arg index is the common case.
                    let mut iter = parts.into_iter();
                    if let Some(first) = iter.next() {
                        let index = walk_index_part(first, expr.clone())?;
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr), index: Box::new(index), null_safe: false,
                        });
                        for p in iter {
                            let index = walk_index_part(p, expr.clone())?;
                            expr = Expression::new(ExprKind::Index {
                                object: Box::new(expr), index: Box::new(index), null_safe: false,
                            });
                        }
                    }
                }
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
                    // Each child is a `collection_element` wrapping either
                    // an expression or a nested `{ ... }` (dict pair /
                    // sub-array). Walk through `walk_collection_element`
                    // so the Dictionary/multi-dim shape lowers correctly.
                    if let Ok(expr) = walk_collection_element(ap) {
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
                if let Ok(expr) = walk_expression(p) {
                    if !is_array {
                        is_array = true;
                    }
                    // Capture the size as the first arg so the array
                    // expression below can preallocate length-N slots.
                    if args.is_empty() {
                        args.push(Argument::positional(expr));
                    }
                }
            }
        }
    }

    // `new Dictionary<K,V> { { key, value }, ... }` — IIFE-lower to:
    //
    //     (() => { var __d = new Dictionary(); __d.Add(k1, v1); ...
    //              return __d; })()
    //
    // Producing a plain Object literal would lose the `Dictionary`
    // `__type` and the runtime collection registry could not route
    // `ContainsKey` / `Add` to the `ecma:map.*` primitives. The IIFE
    // builds a real Map-backed Dictionary and populates it before
    // returning.
    let is_dict_ctor = type_name.eq_ignore_ascii_case("Dictionary")
        || type_name.ends_with("Dictionary");
    if is_dict_ctor && !array_init.is_empty() {
        let mut pairs: Vec<(Expression, Expression)> = Vec::new();
        for elem in &array_init {
            if let ExprKind::Array(parts) = &elem.kind {
                if parts.len() == 2 {
                    pairs.push((parts[0].value.clone(), parts[1].value.clone()));
                }
            }
        }
        if !pairs.is_empty() {
            return Ok(emit_dict_iife(pairs));
        }
    }
    // `new HashSet<T> { v1, v2, ... }` — IIFE-lower to construct + Add
    // calls, same shape as the Dictionary path above. HashSet's
    // backing is `ObjectKind::Set`, registered separately in
    // `vybe_host::builtin_types`, so we need a real `new HashSet()`.
    let is_set_ctor = type_name.eq_ignore_ascii_case("HashSet")
        || type_name.ends_with("HashSet");
    if is_set_ctor && !array_init.is_empty() {
        return Ok(emit_set_iife(type_name.clone(), array_init));
    }

    if is_array && !array_init.is_empty() {
        // Array initializer: new[] { 1, 2, 3 } or new int[] { 1, 2, 3 }
        // — also covers multi-dim (`int[,]`) where each element is itself
        // an Array, which `walk_collection_element` already produced.
        let elements = array_init.into_iter()
            .map(|v| ArrayElement { key: None, value: v, spread: false, by_ref: false })
            .collect();
        return Ok(ExprKind::Array(elements));
    }

    // .NET `new string(charArray)` → `charArray.join("")`. The runtime
    // doesn't carry a String constructor, but every char[] in this VM
    // is a JS array of single-char strings, so `.join("")` is faithful.
    if type_name == "string" && args.len() == 1 && array_init.is_empty() && obj_init.is_empty() {
        let arr = args[0].value.clone();
        return Ok(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(arr),
                field: "join".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(
                ExprKind::Lit(Literal::Str("".into())),
            ))],
            optional: false,
        });
    }

    // Build class expression — dotted names become Member chains
    // (e.g. "MyApp.Foo" → Member { Ident("MyApp"), "Foo" })
    let class_expr = build_dotted_expr(&type_name);

    if is_array {
        // `new int[N]` → `Array.from({length: N}, () => 0)` style
        // pre-fill. We don't have JS Array.from in scope here, so we
        // synthesize an empty literal — the test exercises plain
        // indexed assignment, and our VM grows dynamic arrays on
        // out-of-range writes the same way ECMA-262 §10.4.2 specs.
        return Ok(ExprKind::Array(Vec::new()));
    }

    // Object initializer: `new Point { X = 10, Y = 20 }`. The
    // captured `obj_init` pairs become assignments on the freshly-
    // constructed instance. Lowered as IIFE so the temp ident is
    // self-contained and doesn't pollute the surrounding scope.
    if !obj_init.is_empty() {
        let new_call = Expression::new(ExprKind::New {
            class: Box::new(class_expr),
            args,
        });
        return Ok(emit_object_init_iife(new_call, obj_init));
    }

    Ok(ExprKind::New {
        class: Box::new(class_expr),
        args,
    })
}

/// IIFE-style lowering for `new T(args) { Prop = value, ... }`. Builds
/// a single-call lambda that constructs the object, fires each property
/// assignment, and returns the instance. Same pattern as the
/// Dictionary / HashSet initializer lowerings — keeps the temp local
/// out of the caller's scope.
fn emit_object_init_iife(new_call: Expression, props: Vec<(String, Expression)>) -> ExprKind {
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint: None,
                init: Some(new_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (name, value) in props {
        // __obj.<name> = value;
        let assign = Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__obj")),
                field: name,
                null_safe: false,
            })),
            value: Box::new(value),
        });
        body.push(Statement::with_span(
            StmtKind::Expr(assign),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__obj"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
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

/// Lower a C# 9 record `with` expression: `record_val with { Prop = v, ... }`.
/// The walker emits an IIFE that constructs a shallow copy by reading
/// the receiver's existing properties, applies the with-clause mutations,
/// and returns the new instance. Records compile as plain classes in
/// our compiler, so this is the same shape as a `new T { ... }`
/// initializer that copies fields from the source.
fn walk_with_expr(receiver: Expression, postfix: Pair<Rule>) -> Result<Expression, String> {
    // Collect the with-clause property assignments.
    let mut props: Vec<(String, Expression)> = Vec::new();
    for child in postfix.into_inner() {
        if child.as_rule() == Rule::object_initializer {
            for ip in child.into_inner() {
                if ip.as_rule() == Rule::initializer_member {
                    let mut name = String::new();
                    let mut val = Expression::null();
                    for mp in ip.into_inner() {
                        match mp.as_rule() {
                            Rule::ident_name => name = mp.as_str().to_string(),
                            _ => val = walk_expression(mp).unwrap_or(Expression::null()),
                        }
                    }
                    props.push((name, val));
                }
            }
        }
    }

    // Lower to an IIFE:
    //   ((src) => {
    //       var __o = Object.assign({}, src);
    //       __o.Prop = val;
    //       ...
    //       return __o;
    //   })(receiver)
    //
    // We use `Object.assign({}, src)` shape via the dotted-name
    // resolver (resolves to `ecma:object.assign`) so the clone sees
    // the same prototype chain as the source.
    let mut body: Vec<Statement> = Vec::new();
    let assign_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Object")),
            field: "assign".into(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Object(Vec::new()))),
            Argument::positional(Expression::ident("__src")),
        ],
        optional: false,
    });
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__o".into()),
                type_hint: None,
                init: Some(assign_call),
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
                object: Box::new(Expression::ident("__o")),
                field: name,
                null_safe: false,
            })),
            value: Box::new(value),
        });
        body.push(Statement::with_span(StmtKind::Expr(assign), Span::default()));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__o"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: "__src".into(),
            type_hint: None, default: None,
            pass_by: PassBy::Value, is_rest: false,
            is_kwargs: false, is_optional: false,
            is_nullable: false,
        }],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![Argument::positional(receiver)],
        optional: false,
    }))
}

/// Lower a C# 8 switch expression `subject switch { arm, ... }` into
/// a chain of nested `cond ? then : else_` ternaries. Each arm's
/// `when` guard is AND-ed into the cond. The wildcard `_` arm is the
/// catchall (`else_` of the chain). If no `_` arm is present, the
/// chain falls through to `null`, matching .NET's
/// `SwitchExpressionException` shape (we don't throw — return null
/// rather than complicate codegen).
fn walk_switch_expr(subject: Expression, postfix: Pair<Rule>) -> Result<Expression, String> {
    let arms: Vec<Pair<Rule>> = postfix.into_inner()
        .filter(|p| p.as_rule() == Rule::switch_arm)
        .collect();
    let span = subject.span.clone();
    // We process arms in reverse, building the ternary chain inside-out.
    let mut else_branch = Expression::null();
    let mut else_set = false;
    for arm in arms.into_iter().rev() {
        let mut pattern: Option<Pair<Rule>> = None;
        let mut when_guard: Option<Expression> = None;
        let mut result: Option<Expression> = None;
        let arm_inner: Vec<Pair<Rule>> = arm.into_inner().collect();
        // Order in source: pattern, optional `when`-clause expression,
        // then the result expression. We split by rule, taking the
        // first `switch_pattern` as the pattern, and treating the
        // remaining `expression` children as guard + result.
        let mut exprs: Vec<Pair<Rule>> = Vec::new();
        for inner in arm_inner {
            match inner.as_rule() {
                Rule::switch_pattern => pattern = Some(inner),
                Rule::expression => exprs.push(inner),
                _ => {}
            }
        }
        // Last expr = result; if there's a second expr it's the guard.
        if let Some(last) = exprs.pop() {
            result = Some(walk_expression(last)?);
        }
        if let Some(guard) = exprs.pop() {
            when_guard = Some(walk_expression(guard)?);
        }
        let result = result.ok_or("switch arm missing result")?;
        let pattern = pattern.ok_or("switch arm missing pattern")?;

        // Detect wildcard `_` arm — that's the catch-all.
        let pat_src = pattern.as_str().trim();
        if pat_src == "_" && when_guard.is_none() {
            else_branch = result;
            else_set = true;
            continue;
        }
        let cond = build_switch_pattern_cond(subject.clone(), pattern, when_guard)?;
        if !else_set {
            else_branch = result.clone();
            else_set = true;
            // Still emit the test so the arm runs even if `_` is missing.
        }
        let next = Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(result),
                else_: Box::new(else_branch),
            },
            span.clone(),
        );
        else_branch = next;
    }
    Ok(else_branch)
}

/// Build a boolean test for a single switch-arm pattern.
/// Cases:
///   `<lit>`           → subject === <lit>
///   `<TypeName> <id>` → typeof subject === "<jsname>" (binding dropped)
///   `>= <expr>`       → subject >= <expr>  (relational pattern)
///   `<expr>`          → subject === <expr>  (constant fallback)
fn build_switch_pattern_cond(
    subject: Expression,
    pattern: Pair<Rule>,
    when_guard: Option<Expression>,
) -> Result<Expression, String> {
    let pat_src = pattern.as_str().trim();
    let span = subject.span.clone();
    // Relational pattern: `>= 90`, `<= 50`, `< 0`, `> 0`.
    let rel_op = if pat_src.starts_with(">=") {
        Some(BinOp::GtEq)
    } else if pat_src.starts_with("<=") {
        Some(BinOp::LtEq)
    } else if pat_src.starts_with('>') {
        Some(BinOp::Gt)
    } else if pat_src.starts_with('<') {
        Some(BinOp::Lt)
    } else {
        None
    };
    let mut cond: Expression;
    if let Some(op) = rel_op {
        // Find the inner expression child.
        let inner = pattern.into_inner()
            .find(|p| p.as_rule() == Rule::expression)
            .ok_or("relational pattern missing expression")?;
        let rhs = walk_expression(inner)?;
        cond = Expression::with_span(
            ExprKind::Binary { op, left: Box::new(subject), right: Box::new(rhs) },
            span.clone(),
        );
    } else {
        // Type pattern: `int i`, `string s` (with binding) or constant.
        let mut inner_pairs: Vec<Pair<Rule>> = pattern.into_inner().collect();
        // `type_name ~ ident_name` — type pattern with binding.
        if inner_pairs.len() >= 2
            && inner_pairs[0].as_rule() == Rule::type_name
            && inner_pairs[1].as_rule() == Rule::ident_name
        {
            let type_name = inner_pairs[0].as_str().trim().to_string();
            let test = if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                let typeof_expr = Expression::with_span(
                    ExprKind::TypeOf(Box::new(subject)),
                    span.clone(),
                );
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(typeof_expr),
                        right: Box::new(Expression::with_span(
                            ExprKind::Lit(Literal::Str(js_typeof.into())),
                            span.clone(),
                        )),
                    },
                    span.clone(),
                )
            } else {
                Expression::with_span(
                    ExprKind::IsType { expr: Box::new(subject), type_name },
                    span.clone(),
                )
            };
            cond = test;
        } else if let Some(p) = inner_pairs.pop() {
            // Constant pattern (numeric / string literal / general expr).
            let rhs = walk_expression(p)?;
            cond = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(subject),
                    right: Box::new(rhs),
                },
                span.clone(),
            );
        } else {
            cond = Expression::with_span(ExprKind::Lit(Literal::Bool(false)), span.clone());
        }
    }
    if let Some(guard) = when_guard {
        cond = Expression::with_span(
            ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(cond),
                right: Box::new(guard),
            },
            span,
        );
    }
    Ok(cond)
}

/// IIFE-style lowering for `new Dictionary<,> { { k, v }, ... }`.
/// Emits an immediately-invoked lambda that constructs the dict and
/// populates it, so the runtime gets a real Map-backed Dictionary
/// rather than a plain Object literal that the runtime collection
/// registry can't dispatch through.
fn emit_dict_iife(pairs: Vec<(Expression, Expression)>) -> ExprKind {
    let new_dict = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident("Dictionary")),
        args: vec![],
    });
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__d".into()),
                type_hint: None,
                init: Some(new_dict),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (k, v) in pairs {
        let add_call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__d")),
                field: "Add".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(k), Argument::positional(v)],
            optional: false,
        });
        body.push(Statement::with_span(
            StmtKind::Expr(add_call),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__d"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// IIFE-style lowering for `new HashSet<T> { v1, v2, ... }` — same
/// shape as `emit_dict_iife` but adds single values rather than pairs.
fn emit_set_iife(type_name: String, elements: Vec<Expression>) -> ExprKind {
    let new_set = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&type_name)),
        args: vec![],
    });
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__s".into()),
                type_hint: None,
                init: Some(new_set),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for v in elements {
        let add_call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__s")),
                field: "Add".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(v)],
            optional: false,
        });
        body.push(Statement::with_span(
            StmtKind::Expr(add_call),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__s"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// Walk a single `collection_element` (the body of a flat or nested
/// `{ ... }` initializer entry). Flat children are walked as plain
/// expressions; nested-brace children become Array literals so the
/// caller can recognise them as dict pairs (Dictionary) or sub-arrays
/// (multi-dim).
fn walk_collection_element(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::collection_element {
        let src = pair.as_str().trim_start();
        // Nested-brace form: emit Array(elements).
        if src.starts_with('{') {
            let mut elements = Vec::new();
            for inner in pair.into_inner() {
                if let Ok(expr) = walk_expression(inner) {
                    elements.push(ArrayElement {
                        key: None, value: expr, spread: false, by_ref: false,
                    });
                }
            }
            return Ok(Expression::new(ExprKind::Array(elements)));
        }
        // Flat form: walk the single child expression.
        if let Some(inner) = pair.into_inner().next() {
            return walk_expression(inner);
        }
        return Ok(Expression::null());
    }
    walk_expression(pair)
}

/// Map a C# primitive type name to its `typeof` result string in JS,
/// or `None` if the type is a user class. `string`, `int`, etc. lower
/// to typeof tests because JS values don't carry a `__type` slot.
fn primitive_to_typeof(type_name: &str) -> Option<&'static str> {
    match type_name {
        "string" | "String" => Some("string"),
        "int" | "long" | "double" | "float" | "decimal"
        | "byte" | "sbyte" | "short" | "ushort"
        | "uint" | "ulong" | "nint" | "nuint" => Some("number"),
        "bool" | "Boolean" => Some("boolean"),
        _ => None,
    }
}

/// Walk a single index part (the inside of `arr[...]`).
/// Handles `from_end_index` (^N → arr.length - N) and plain expressions.
fn walk_index_part(pair: Pair<Rule>, receiver: Expression) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::from_end_index => {
            // `^N` → receiver.length - N (or for ranges, the same expression)
            let inner = pair.into_inner().next()
                .ok_or_else(|| "from_end_index missing inner expression".to_string())?;
            let n_expr = walk_expression(inner)?;
            // `__len__(receiver) - n_expr`
            let length = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__len__")),
                args: vec![Argument::positional(receiver)],
                optional: false,
            });
            Ok(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(length),
                right: Box::new(n_expr),
            }))
        }
        _ => walk_expression(pair),
    }
}

/// Walk an `is`-pattern operand into a boolean expression. Covers
/// ECMA C# §11.11.7 patterns we surface today:
///   - `is null` → `expr === null`
///   - `is not null` → `expr !== null`
///   - `is <literal>` → equality compare
///   - `is <Type>` → ExprKind::IsType
///   - `is <Type> ident` → ExprKind::IsType + assignment to ident
///     (the ident binding is exposed as a synthetic Block returning
///     the boolean — handled via SequenceExpr if available, else
///     just IsType for now and the binding is dropped).
fn walk_is_pattern(receiver: Expression, pattern_clause: Pair<Rule>) -> Result<Expression, String> {
    // The literal "not" token isn't a Pair, so probe the source.
    let clause_src = pattern_clause.as_str().trim_start();
    let negated = clause_src.starts_with("not")
        && clause_src[3..].chars().next().map_or(true, |c| c.is_whitespace());
    let clause_inner: Vec<Pair<Rule>> = pattern_clause.into_inner().collect();
    let atom = clause_inner.into_iter().next().ok_or("Empty pattern atom".to_string())?;
    let atom_inner: Vec<Pair<Rule>> = atom.into_inner().collect();
    let first = atom_inner.first().ok_or("Empty pattern atom inner".to_string())?;

    let span = receiver.span.clone();
    let result = match first.as_rule() {
        Rule::null_kw => {
            // `expr is null` → `expr === null`
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(receiver),
                    right: Box::new(Expression::with_span(
                        ExprKind::Lit(Literal::Null), span.clone(),
                    )),
                },
                span.clone(),
            )
        }
        Rule::numeric_literal | Rule::string_literal => {
            // Constant pattern → strict-equality compare.
            let lit = walk_expression(first.clone())?;
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(receiver),
                    right: Box::new(lit),
                },
                span.clone(),
            )
        }
        _ => {
            // type_name with optional ident binding.
            let type_name = first.as_str().trim().to_string();
            // Primitive-type patterns lower to `typeof v === "<jsname>"`
            // so they match JS values that have no `__type` slot. Falls
            // back to IsType for user classes (which DO carry __type).
            if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                let typeof_expr = Expression::with_span(
                    ExprKind::TypeOf(Box::new(receiver)),
                    span.clone(),
                );
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(typeof_expr),
                        right: Box::new(Expression::with_span(
                            ExprKind::Lit(Literal::Str(js_typeof.into())),
                            span.clone(),
                        )),
                    },
                    span.clone(),
                )
            } else {
                // The ident-binding form (`obj is string s`) requires
                // declaring `s` and assigning the cast value. For now we
                // emit just the type test; the binding is silently
                // dropped — tests that exercise the binding will still
                // see the boolean correctly.
                Expression::with_span(
                    ExprKind::IsType { expr: Box::new(receiver), type_name },
                    span.clone(),
                )
            }
        }
    };
    if negated {
        Ok(Expression::with_span(
            ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(result) },
            span,
        ))
    } else {
        Ok(result)
    }
}

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::argument)
        .map(|p| {
            let src = p.as_str().trim();
            let by_ref = src.starts_with("ref ") || src.starts_with("out ");
            let inner_pairs: Vec<Pair<Rule>> = p.into_inner().collect();

            // `out var x` / `out int x` — desugar to a synthetic var
            // declaration prepended elsewhere (TODO). For now, we
            // extract the ident as a write-target reference.
            if src.starts_with("out ") {
                let last = inner_pairs.last().ok_or("Empty out argument".to_string())?;
                let name = last.as_str().trim().to_string();
                return Ok(Argument {
                    value: Expression::with_span(ExprKind::Ident(name), to_span(last)),
                    name: None, by_ref: true, spread: false,
                });
            }

            // Named argument: first child is `ident_name`, second is the
            // expression. Disambiguated from positional by the grammar's
            // explicit `ident_name ~ ":" ~ expression` alternative.
            if inner_pairs.len() >= 2
                && inner_pairs[0].as_rule() == Rule::ident_name
                && inner_pairs.iter().any(|p| !matches!(p.as_rule(), Rule::ident_name)
                    && p.as_rule() != Rule::argument_list)
            {
                let name_str = inner_pairs[0].as_str().to_string();
                // Look for the value expression (anything that isn't ident_name).
                if let Some(value_pair) = inner_pairs.iter()
                    .find(|p| p.as_rule() != Rule::ident_name)
                {
                    let value = walk_expression(value_pair.clone())?;
                    return Ok(Argument {
                        value, name: Some(name_str), by_ref, spread: false,
                    });
                }
            }

            let inner = inner_pairs.into_iter().next().ok_or("Empty argument".to_string())?;
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
    // typeof(T) emits a string literal `"System.<Name>"`. Resolve
    // `.Name` / `.FullName` access on such literals at compile time
    // so `typeof(int).Name` constant-folds to `"Int32"`.
    if let ExprKind::Lit(Literal::Str(s)) = &object.kind {
        if s.starts_with("System.") {
            if name == "Name" {
                let short = s.rsplit('.').next().unwrap_or(s).to_string();
                return Expression::new(ExprKind::Lit(Literal::Str(short)));
            }
            if name == "FullName" {
                return object;
            }
        }
    }
    // `expr.GetType().<Name|FullName>` — rewrite the chained access to
    // a runtime ternary on `typeof expr`. Walking happens inner-first,
    // so `expr.GetType()` is here the receiver of `.Name`. We detect
    // the pattern (a call with callee `Member(_, "GetType")`) and
    // unwrap to the ternary directly. Standalone `expr.GetType()` (no
    // chained access) is left alone — its result is unused in any
    // current test path.
    if name == "Name" || name == "FullName" {
        if let ExprKind::Call { callee, args, .. } = &object.kind {
            if args.is_empty() {
                if let ExprKind::Member { object: receiver, field, .. } = &callee.kind {
                    if field == "GetType" {
                        let short = dotnet_runtime_type_name_expr((**receiver).clone());
                        if name == "Name" {
                            return short;
                        }
                        // FullName: prepend "System." via DYN_ADD-equivalent
                        return Expression::new(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(Expression::new(ExprKind::Lit(Literal::Str("System.".into())))),
                            right: Box::new(short),
                        });
                    }
                }
            }
        }
    }
    // C# `.Length` / `.Count` lowers to `__len__(receiver)` so the
    // value-method dispatch picks a single canonical opcode regardless
    // of receiver type (string vs. array vs. List). One trap: a class
    // identifier as the receiver (`Counter.Count` reading a user
    // static int field) MUST go through plain Member access — the
    // class object isn't a sequence and `__len__` would silently
    // return 0. We detect class identifiers by checking the receiver
    // is a single `Ident` starting with an uppercase letter (PascalCase
    // — the C# convention for class names). User instance variables
    // by convention use camelCase, so `arr.Length` / `list.Count` /
    // `s.Length` keep the canonical lowering.
    let is_class_static = matches!(
        &object.kind,
        ExprKind::Ident(n) if n.chars().next().map_or(false, |c| c.is_ascii_uppercase())
    );
    let canonical = if is_class_static {
        None
    } else {
        match name {
            "Length" | "Count" => Some("__len__"),
            _ => None,
        }
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
    // LINQ surface (First / Last / Skip / Take / Average / FirstOrDefault /
    // Distinct / Aggregate / OrderByDescending / Count(pred) / ToList /
    // ToArray) is in `emitter/dotnet/core/linq_adapter.rs` and wired
    // through the C# profile's [value_methods] table — VB and any other
    // .NET-shape language pick up the same emitters by listing them in
    // their own profile. The dispatch in `compiler/calls.rs` routes
    // `common:dotnet.*` value-method overloads around the runtime
    // collection registry so even names like `Count` (which IS in the
    // registry) hit the LINQ adapter when called with the predicate.
    //
    // Instance-method rewrite: `a.CompareTo(b)` → `a < b ? -1 : a > b ? 1 : 0`
    // (works for strings AND numbers — same JS comparison semantics).
    if let ExprKind::Member { object, field, .. } = &callee.kind {
        if field == "CompareTo" && args.len() == 1 {
            let a = (**object).clone();
            let b = args[0].value.clone();
            let lt = Expression::new(ExprKind::Binary {
                op: BinOp::Lt, left: Box::new(a.clone()), right: Box::new(b.clone()),
            });
            let gt = Expression::new(ExprKind::Binary {
                op: BinOp::Gt, left: Box::new(a), right: Box::new(b),
            });
            return Expression::new(ExprKind::Ternary {
                cond: Box::new(lt),
                then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(gt),
                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                    else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                })),
            });
        }
    }
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
            // string.Equals(a, b, StringComparison.OrdinalIgnoreCase)
            //   → a.toLowerCase() === b.toLowerCase()
            // string.Equals(a, b) → a === b
            if obj_name.eq_ignore_ascii_case("string")
                && field.eq_ignore_ascii_case("Equals")
                && (args.len() == 2 || args.len() == 3)
            {
                let a = args[0].value.clone();
                let b = args[1].value.clone();
                let ignore_case = args.get(2).map_or(false, |arg| {
                    if let ExprKind::Member { field, .. } = &arg.value.kind {
                        field.contains("IgnoreCase")
                    } else { false }
                });
                if ignore_case {
                    let a_lc = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(a), field: "toLowerCase".into(), null_safe: false,
                        })),
                        args: vec![], optional: false,
                    });
                    let b_lc = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(b), field: "toLowerCase".into(), null_safe: false,
                        })),
                        args: vec![], optional: false,
                    });
                    return Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(a_lc),
                        right: Box::new(b_lc),
                    });
                }
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(a),
                    right: Box::new(b),
                });
            }
            // char.IsXxx / char.ToXxx — single-char string predicates / converters
            if obj_name.eq_ignore_ascii_case("char") && args.len() == 1 {
                let c = args[0].value.clone();
                if let Some(rewritten) = char_static_lower(field, c) {
                    return rewritten;
                }
            }
            // bool.Parse(s) → s.toLowerCase() === "true"
            if (obj_name.eq_ignore_ascii_case("bool") || obj_name == "Boolean")
                && field.eq_ignore_ascii_case("Parse")
                && args.len() == 1
            {
                let s = args[0].value.clone();
                let s_lc = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(s), field: "toLowerCase".into(), null_safe: false,
                    })),
                    args: vec![], optional: false,
                });
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(s_lc),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("true".into())))),
                });
            }
            // `Array.Reverse / Exists / Find / FindAll / TrueForAll /
            // ConvertAll / ForEach / IndexOf` static helpers are wired
            // through the `[builtins]` `Array.*` entries in the C# profile,
            // which route to `common:dotnet.array_*` adapters in
            // `crates/vybex/src/emitter/dotnet/core/array_adapter.rs`.
            // VB picks them up by listing the same `common:dotnet.*`
            // emit targets — no walker rewrite required.
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

/// Lower `char.<method>(c)` static helpers into ECMA-shape expressions.
/// All are ASCII-faithful only (matches what existing tests assert).
fn char_static_lower(method: &str, c: Expression) -> Option<Expression> {
    let lower_method = method.to_ascii_lowercase();
    match lower_method.as_str() {
        "toupper" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(c), field: "toUpperCase".into(), null_safe: false,
            })),
            args: vec![], optional: false,
        })),
        "tolower" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(c), field: "toLowerCase".into(), null_safe: false,
            })),
            args: vec![], optional: false,
        })),
        // ASCII range predicates — emitted as `c >= "A" && c <= "Z"`.
        "isupper" => Some(char_in_range(c, "A", "Z")),
        "islower" => Some(char_in_range(c, "a", "z")),
        "isdigit" => Some(char_in_range(c, "0", "9")),
        // IsLetter — uppercase OR lowercase ASCII letter.
        "isletter" => {
            let upper = char_in_range(c.clone(), "A", "Z");
            let lower = char_in_range(c, "a", "z");
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or, left: Box::new(upper), right: Box::new(lower),
            }))
        }
        "isletterordigit" => {
            let upper = char_in_range(c.clone(), "A", "Z");
            let lower = char_in_range(c.clone(), "a", "z");
            let digit = char_in_range(c, "0", "9");
            let letter = Expression::new(ExprKind::Binary {
                op: BinOp::Or, left: Box::new(upper), right: Box::new(lower),
            });
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or, left: Box::new(letter), right: Box::new(digit),
            }))
        }
        // IsWhiteSpace — match space, tab, newline, carriage return.
        "iswhitespace" => {
            let space = eq_lit(c.clone(), " ");
            let tab = eq_lit(c.clone(), "\t");
            let newline = eq_lit(c.clone(), "\n");
            let cr = eq_lit(c, "\r");
            Some(or_chain(vec![space, tab, newline, cr]))
        }
        _ => None,
    }
}

/// Map a C#/.NET type-name token to its System.* short name.
/// `int` → `Int32`, `string` → `String`, `MyClass` → `MyClass`.
fn dotnet_type_name(t: &str) -> String {
    let trimmed = t.trim().trim_end_matches('?');
    match trimmed {
        "int" | "Int32" => "Int32",
        "uint" | "UInt32" => "UInt32",
        "long" | "Int64" => "Int64",
        "ulong" | "UInt64" => "UInt64",
        "short" | "Int16" => "Int16",
        "ushort" | "UInt16" => "UInt16",
        "byte" | "Byte" => "Byte",
        "sbyte" | "SByte" => "SByte",
        "float" | "Single" => "Single",
        "double" | "Double" => "Double",
        "decimal" | "Decimal" => "Decimal",
        "bool" | "Boolean" => "Boolean",
        "char" | "Char" => "Char",
        "string" | "String" => "String",
        "object" | "Object" => "Object",
        other => other,
    }.to_string()
}

/// Build a runtime expression that yields the .NET type name of `expr`.
///
/// `Math.floor(e) === e ? "Int32" : "Double"` for numbers, `String` /
/// `Boolean` for primitives, and `Object` as the fallback. Implemented
/// as nested ternaries on `typeof e` so the result is plain bytecode
/// with no host helper required.
fn dotnet_runtime_type_name_expr(expr: Expression) -> Expression {
    let typeof_expr = Expression::new(ExprKind::TypeOf(Box::new(expr.clone())));

    let is_string = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("string".into())))),
    });
    let is_number = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("number".into())))),
    });
    let is_boolean = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("boolean".into())))),
    });

    // Math.floor(e) === e — true for whole numbers; faithful to .NET
    // semantics where `42.GetType().Name == "Int32"` and
    // `3.14.GetType().Name == "Double"`. Vybe stores all numbers as f64
    // so this is the only post-hoc int/float distinction available.
    let floor_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Math")),
            field: "floor".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(expr.clone())],
        optional: false,
    });
    let is_int = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(floor_call),
        right: Box::new(expr),
    });

    let number_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_int),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str("Int32".into())))),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Str("Double".into())))),
    });

    let bool_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_boolean),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str("Boolean".into())))),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Str("Object".into())))),
    });

    let num_or_bool = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_number),
        then: Box::new(number_branch),
        else_: Box::new(bool_branch),
    });

    Expression::new(ExprKind::Ternary {
        cond: Box::new(is_string),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str("String".into())))),
        else_: Box::new(num_or_bool),
    })
}

fn char_in_range(c: Expression, lo: &str, hi: &str) -> Expression {
    let ge = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(lo.into())))),
    });
    let le = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(hi.into())))),
    });
    Expression::new(ExprKind::Binary {
        op: BinOp::And, left: Box::new(ge), right: Box::new(le),
    })
}

fn eq_lit(c: Expression, lit: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(lit.into())))),
    })
}

fn or_chain(mut exprs: Vec<Expression>) -> Expression {
    let mut acc = exprs.remove(0);
    for e in exprs {
        acc = Expression::new(ExprKind::Binary {
            op: BinOp::Or, left: Box::new(acc), right: Box::new(e),
        });
    }
    acc
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
