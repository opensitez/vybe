use pest::Parser;
use pest::iterators::Pair;
use super::{PascalParser, Rule};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs = PascalParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut name = "main".to_string();

    for pair in pairs {
        if pair.as_rule() != Rule::program { continue; }
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::program_heading => {
                    // program Foo; or unit Foo;
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::identifier {
                            name = p.as_str().to_string();
                        }
                    }
                }
                Rule::uses_clause => {
                    let span = to_span(&inner);
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::identifier {
                            imports.push(Import {
                                kind: ImportKind::Simple {
                                    path: p.as_str().to_string(),
                                    alias: None,
                                },
                                span,
                            });
                        }
                    }
                }
                Rule::interface_section | Rule::implementation_section => {
                    // Markers only — no content to walk
                }
                Rule::decl_section => {
                    walk_decl_section(inner, &mut body)?;
                }
                Rule::program_body => {
                    // compound_statement wrapping main body
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::compound_statement {
                            let stmts = walk_compound_statement(p)?;
                            body.extend(stmts);
                        }
                    }
                }
                Rule::EOI => {}
                _ => {}
            }
        }
    }

    // Pascal allows method bodies to be implemented outside the class declaration
    // (e.g. `constructor TFoo.Create(...) begin ... end;`). Merge those standalone
    // FunctionDecls back into the matching ClassDecl so the compiler sees them as
    // ordinary class members.
    merge_separated_methods(&mut body);

    // Now that class declarations are stable, rewrite `TFoo.Create(args)` (Pascal's
    // constructor invocation syntax) into the canonical `New { class: TFoo, args }`
    // AST so every language ends up with the same instantiation node.
    let class_names: std::collections::HashSet<String> = body.iter().filter_map(|s| {
        if let StmtKind::ClassDecl { name, .. } = &s.kind {
            Some(name.to_lowercase())
        } else { None }
    }).collect();
    for stmt in body.iter_mut() {
        rewrite_constructor_calls_stmt(stmt, &class_names);
    }

    Ok(Module {
        name,
        language: Lang::Pascal,
        body,
        imports,
    })
}

/// Walk a statement and rewrite `ClassName.Create(args)` into `New { class, args }`
/// when `ClassName` matches a class declared in the same module.
fn rewrite_constructor_calls_stmt(stmt: &mut Statement, classes: &std::collections::HashSet<String>) {
    match &mut stmt.kind {
        StmtKind::Expr(e) => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Block(stmts) => for s in stmts { rewrite_constructor_calls_stmt(s, classes); },
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &mut d.init { rewrite_constructor_calls_expr(e, classes); }
            }
        }
        StmtKind::FunctionDecl { body, .. } => for s in body { rewrite_constructor_calls_stmt(s, classes); },
        StmtKind::ClassDecl { members, .. } => {
            for m in members { rewrite_constructor_calls_member(m, classes); }
        }
        StmtKind::StructDecl { members, .. } | StmtKind::ModuleDecl { members, .. } => {
            for m in members { rewrite_constructor_calls_member(m, classes); }
        }
        StmtKind::NamespaceDecl { body, .. } => for s in body { rewrite_constructor_calls_stmt(s, classes); },
        StmtKind::If { cond, then_body, elifs, else_body } => {
            rewrite_constructor_calls_expr(cond, classes);
            for s in then_body { rewrite_constructor_calls_stmt(s, classes); }
            for (c, b) in elifs {
                rewrite_constructor_calls_expr(c, classes);
                for s in b { rewrite_constructor_calls_stmt(s, classes); }
            }
            if let Some(b) = else_body { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
        }
        StmtKind::For { init, cond, update, body } => {
            if let Some(i) = init { rewrite_constructor_calls_stmt(i, classes); }
            if let Some(c) = cond { rewrite_constructor_calls_expr(c, classes); }
            if let Some(u) = update { rewrite_constructor_calls_expr(u, classes); }
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
        }
        StmtKind::ForIn { iter, body, else_body, .. } => {
            rewrite_constructor_calls_expr(iter, classes);
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
            if let Some(b) = else_body { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
        }
        StmtKind::While { cond, body, else_body } => {
            rewrite_constructor_calls_expr(cond, classes);
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
            if let Some(b) = else_body { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
            rewrite_constructor_calls_expr(cond, classes);
        }
        StmtKind::Switch { expr, cases, default } => {
            rewrite_constructor_calls_expr(expr, classes);
            for c in cases { for s in &mut c.body { rewrite_constructor_calls_stmt(s, classes); } }
            if let Some(b) = default { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
            for c in catches { for s in &mut c.body { rewrite_constructor_calls_stmt(s, classes); } }
            if let Some(b) = else_body { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
            if let Some(b) = finally { for s in b { rewrite_constructor_calls_stmt(s, classes); } }
        }
        StmtKind::With { items, body, .. } => {
            for it in items { rewrite_constructor_calls_expr(&mut it.expr, classes); }
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
        }
        StmtKind::Return(Some(e)) => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Throw { expr: Some(e), .. } => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Assign { targets, value } => {
            for t in targets { rewrite_constructor_calls_expr(t, classes); }
            rewrite_constructor_calls_expr(value, classes);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_constructor_calls_expr(target, classes);
            rewrite_constructor_calls_expr(value, classes);
        }
        _ => {}
    }
}

fn rewrite_constructor_calls_member(m: &mut ClassMember, classes: &std::collections::HashSet<String>) {
    match m {
        ClassMember::Field { init: Some(e), .. } => rewrite_constructor_calls_expr(e, classes),
        ClassMember::Method(stmt) => rewrite_constructor_calls_stmt(stmt, classes),
        ClassMember::Constructor { body, .. } => {
            for s in body { rewrite_constructor_calls_stmt(s, classes); }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(g) = getter { for s in g { rewrite_constructor_calls_stmt(s, classes); } }
            if let Some(set) = setter { for s in &mut set.body { rewrite_constructor_calls_stmt(s, classes); } }
        }
        ClassMember::Const { value, .. } => rewrite_constructor_calls_expr(value, classes),
        ClassMember::NestedType(stmt) => rewrite_constructor_calls_stmt(stmt, classes),
        _ => {}
    }
}

fn rewrite_constructor_calls_expr(expr: &mut Expression, classes: &std::collections::HashSet<String>) {
    // Check Call(Member(ClassName, "Create"), args) BEFORE descending so the
    // Member-only rewrite below doesn't fire on the callee position first and
    // turn `TFoo.Create(42)` into a call on a New expression.
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                if classes.contains(&class_name.to_lowercase())
                   && field.eq_ignore_ascii_case("Create")
                {
                    let new_class = Box::new(Expression::ident(class_name));
                    let mut new_args = args.clone();
                    for a in new_args.iter_mut() { rewrite_constructor_calls_expr(&mut a.value, classes); }
                    expr.kind = ExprKind::New { class: new_class, args: new_args };
                    return;
                }
            }
        }
    }

    // First descend into children, then check this node so deeply-nested
    // patterns are also normalized.
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_constructor_calls_expr(left, classes);
            rewrite_constructor_calls_expr(right, classes);
        }
        ExprKind::Unary { expr: e, .. } => rewrite_constructor_calls_expr(e, classes),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_constructor_calls_expr(cond, classes);
            rewrite_constructor_calls_expr(then, classes);
            rewrite_constructor_calls_expr(else_, classes);
        }
        ExprKind::Member { object, .. } => rewrite_constructor_calls_expr(object, classes),
        ExprKind::Index { object, index } => {
            rewrite_constructor_calls_expr(object, classes);
            rewrite_constructor_calls_expr(index, classes);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_constructor_calls_expr(callee, classes);
            for a in args.iter_mut() { rewrite_constructor_calls_expr(&mut a.value, classes); }
        }
        ExprKind::New { class, args } => {
            rewrite_constructor_calls_expr(class, classes);
            for a in args.iter_mut() { rewrite_constructor_calls_expr(&mut a.value, classes); }
        }
        ExprKind::Assign { target, value } => {
            rewrite_constructor_calls_expr(target, classes);
            rewrite_constructor_calls_expr(value, classes);
        }
        ExprKind::Array(elems) => for el in elems {
            rewrite_constructor_calls_expr(&mut el.value, classes);
        },
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for e in items { rewrite_constructor_calls_expr(e, classes); }
        }
        ExprKind::Object(props) => {
            for p in props {
                if let ObjectProperty::KeyValue { value, .. } = p {
                    rewrite_constructor_calls_expr(value, classes);
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for p in parts {
                match p {
                    InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => rewrite_constructor_calls_expr(e, classes),
                    _ => {}
                }
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_constructor_calls_expr(left, classes);
            rewrite_constructor_calls_expr(right, classes);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_constructor_calls_expr(start, classes);
            rewrite_constructor_calls_expr(end, classes);
        }
        ExprKind::IsType { expr: e, .. } | ExprKind::Cast { expr: e, .. } => rewrite_constructor_calls_expr(e, classes),
        _ => {}
    }

    // Pascal allows zero-arg constructor calls without parens: `f := TFoo.Create;`
    // Detect bare `ClassName.Create` member access on a known class and rewrite
    // it to a zero-arg `New { class, [] }`.
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        if let ExprKind::Ident(class_name) = &object.kind {
            if classes.contains(&class_name.to_lowercase())
               && field.eq_ignore_ascii_case("Create")
            {
                expr.kind = ExprKind::New {
                    class: Box::new(Expression::ident(class_name)),
                    args: Vec::new(),
                };
            }
        }
    }
}

/// Pascal post-processing: attach `ClassName.Method` standalone FunctionDecls
/// to their matching ClassDecl members. The constructor (`ClassName.Create`)
/// fills in `ClassMember::Constructor`. Other methods fill the body of the
/// matching `ClassMember::Method`.
fn merge_separated_methods(body: &mut Vec<Statement>) {
    use std::collections::HashMap;

    // Collect class indices by canonicalized (lowercase) name
    let mut class_idx: HashMap<String, usize> = HashMap::new();
    for (i, s) in body.iter().enumerate() {
        if let StmtKind::ClassDecl { name, .. } = &s.kind {
            class_idx.insert(name.to_lowercase(), i);
        }
    }

    // Walk in reverse so removals don't shift earlier indices
    let mut to_remove: Vec<usize> = Vec::new();
    for i in 0..body.len() {
        let (class_name, method_name, params, ret, mods, body_stmts, is_sub) = {
            let stmt = &body[i];
            let StmtKind::FunctionDecl { name, params, return_type, body: b, modifiers, is_sub, .. } = &stmt.kind else { continue };
            let Some((cls, mth)) = name.split_once('.') else { continue };
            (cls.to_string(), mth.to_string(), params.clone(), return_type.clone(), modifiers.clone(), b.clone(), *is_sub)
        };

        let Some(&ci) = class_idx.get(&class_name.to_lowercase()) else { continue };
        let StmtKind::ClassDecl { members, .. } = &mut body[ci].kind else { continue };

        // Try constructor first: any ClassMember::Constructor whose params arity matches,
        // when the method name is "Create" (Pascal convention) — fall back to first ctor.
        let is_create = method_name.eq_ignore_ascii_case("Create");
        let mut attached = false;
        if is_create {
            for m in members.iter_mut() {
                if let ClassMember::Constructor { params: cp, body: cb, base_args: ba, .. } = m {
                    if cb.is_empty() {
                        *cp = params.clone();
                        let mut new_body = body_stmts.clone();
                        // Pascal pattern: `inherited Create(args)` as the FIRST statement
                        // is the base-constructor invocation. Lift it into `base_args`
                        // so the compiler runs the canonical C#-style path
                        // (parent ctor → field inits → method bindings → body).
                        // This keeps the AST uniform across languages.
                        if let Some(first) = new_body.first() {
                            let extracted = match &first.kind {
                                StmtKind::Expr(e) => match &e.kind {
                                    ExprKind::SuperCall { method, args } => {
                                        let is_ctor = method.is_none()
                                            || method.as_ref().map_or(false, |m| m.eq_ignore_ascii_case("Create"));
                                        if is_ctor {
                                            Some(args.iter().map(|a| a.value.clone()).collect::<Vec<_>>())
                                        } else { None }
                                    }
                                    _ => None,
                                },
                                _ => None,
                            };
                            if let Some(extracted_args) = extracted {
                                *ba = Some(extracted_args);
                                new_body.remove(0);
                            }
                        }
                        *cb = new_body;
                        attached = true;
                        break;
                    }
                }
            }
        }

        if !attached {
            // Find a Method with matching name and empty body
            for m in members.iter_mut() {
                if let ClassMember::Method(stmt) = m {
                    if let StmtKind::FunctionDecl { name: mn, params: mp, body: mb, return_type: mr, modifiers: mm, is_sub: ms, .. } = &mut stmt.kind {
                        if mn.eq_ignore_ascii_case(&method_name) && mb.is_empty() {
                            *mp = params.clone();
                            *mb = body_stmts.clone();
                            *mr = ret.clone();
                            *mm = mods.clone();
                            *ms = is_sub;
                            attached = true;
                            break;
                        }
                    }
                }
            }
        }

        if attached {
            to_remove.push(i);
        }
    }

    for i in to_remove.into_iter().rev() {
        body.remove(i);
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Declaration section
// ════════════════════════════════════════════════════════════════════════════

fn walk_decl_section(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for decl in pair.into_inner() {
        match decl.as_rule() {
            Rule::var_section => {
                body.extend(walk_var_section(decl)?);
            }
            Rule::const_section => {
                body.extend(walk_const_section(decl)?);
            }
            Rule::type_section => {
                body.extend(walk_type_section(decl)?);
            }
            Rule::procedure_decl_or_method => {
                body.push(walk_procedure_decl_or_method(decl)?);
            }
            Rule::function_decl_or_method => {
                body.push(walk_function_decl_or_method(decl)?);
            }
            Rule::constructor_method_impl => {
                body.push(walk_constructor_method_impl(decl)?);
            }
            Rule::destructor_method_impl => {
                body.push(walk_destructor_method_impl(decl)?);
            }
            _ => {}
        }
    }
    Ok(())
}

// ── Var section ────────────────────────────────────────────────────────────

fn walk_var_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::var_decl {
            let decls = walk_var_decl(p)?;
            stmts.push(Statement::with_span(
                StmtKind::VarDecl { declarations: decls, kind: VarDeclKind::Dim },
                span,
            ));
        }
    }
    Ok(stmts)
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<Vec<VarDeclarator>, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => {
                type_hint = Some(type_ref_to_string(&p));
            }
            Rule::var_init => {
                for inner in p.into_inner() {
                    if inner.as_rule() == Rule::expression {
                        init = Some(walk_expression(inner)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(names.into_iter().map(|n| VarDeclarator {
        pattern: BindingPattern::Ident(n),
        type_hint: type_hint.clone(),
        init: init.clone(),
        array_bounds: None,
        with_events: false,
    }).collect())
}

// ── Const section ──────────────────────────────────────────────────────────

fn walk_const_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::const_decl {
            let decl = walk_const_decl(p)?;
            stmts.push(Statement::with_span(
                StmtKind::VarDecl { declarations: vec![decl], kind: VarDeclKind::Const },
                span,
            ));
        }
    }
    Ok(stmts)
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::expression => init = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint,
        init,
        array_bounds: None,
        with_events: false,
    })
}

// ── Type section ───────────────────────────────────────────────────────────

fn walk_type_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::type_decl {
            stmts.push(walk_type_decl(p)?);
        }
    }
    Ok(stmts)
}

fn walk_type_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut type_def_pair: Option<Pair<Rule>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_def => type_def_pair = Some(p),
            _ => {}
        }
    }

    let def = type_def_pair.ok_or("Missing type_def in type_decl")?;
    let inner = def.into_inner().next().ok_or("Empty type_def")?;

    match inner.as_rule() {
        Rule::class_type => walk_class_type(inner, &name, span),
        Rule::record_type => walk_record_type(inner, &name, span),
        Rule::interface_type => walk_interface_type(inner, &name, span),
        Rule::enum_type => walk_enum_type(inner, &name, span),
        Rule::array_type => {
            // Type alias for array: type TMyArray = array[0..9] of Integer;
            // Emit as a VarDecl with type hint
            Ok(Statement::with_span(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: Some(type_ref_to_string(&inner)),
                    init: None,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }, span))
        }
        Rule::pointer_type => {
            // Type alias for pointer
            let target = inner.into_inner().next()
                .map(|p| type_ref_to_string(&p))
                .unwrap_or_default();
            Ok(Statement::with_span(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: Some(format!("^{}", target)),
                    init: None,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }, span))
        }
        Rule::type_alias => {
            // Simple type alias: type TMyInt = Integer;
            let aliased = inner.into_inner().next()
                .map(|p| type_ref_to_string(&p))
                .unwrap_or_default();
            Ok(Statement::with_span(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: Some(aliased),
                    init: None,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }, span))
        }
        other => Err(format!("Unexpected type_def inner: {:?}", other)),
    }
}

// ── Class type ─────────────────────────────────────────────────────────────

fn walk_class_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_heritage => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        parents.push(id.as_str().to_string());
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::field_decl => {
                            members.extend(walk_field_decl_members(m)?);
                        }
                        Rule::class_constructor => {
                            members.push(walk_class_constructor_sig(m)?);
                        }
                        Rule::class_destructor => {
                            members.push(walk_class_method_sig(m, true)?);
                        }
                        Rule::class_procedure => {
                            members.push(walk_class_method_sig(m, false)?);
                        }
                        Rule::class_function => {
                            members.push(walk_class_method_sig(m, false)?);
                        }
                        Rule::class_class_member => {
                            members.push(walk_class_class_member(m)?);
                        }
                        Rule::class_property_decl => {
                            members.push(walk_class_property_decl(m)?);
                        }
                        _ => {
                            // visibility markers are silent (_) rules, skip
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::ClassDecl {
        name: name.to_string(),
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
    }, span))
}

// ── Record type ────────────────────────────────────────────────────────────

fn walk_record_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut members = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::record_body {
            for m in p.into_inner() {
                match m.as_rule() {
                    Rule::field_decl => {
                        members.extend(walk_field_decl_members(m)?);
                    }
                    Rule::record_method_sig => {
                        members.push(walk_record_method_sig(m)?);
                    }
                    Rule::record_class_method => {
                        members.push(walk_record_class_method(m)?);
                    }
                    Rule::record_operator_method => {
                        members.push(walk_record_operator_method(m)?);
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(Statement::with_span(StmtKind::StructDecl {
        name: name.to_string(),
        interfaces: Vec::new(),
        members,
        visibility: Visibility::Public,
    }, span))
}

// ── Interface type ─────────────────────────────────────────────────────────

fn walk_interface_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::interface_heritage => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        parents.push(id.as_str().to_string());
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::interface_procedure => {
                            members.push(walk_interface_method(m, true)?);
                        }
                        Rule::interface_function => {
                            members.push(walk_interface_method(m, false)?);
                        }
                        Rule::interface_property_decl => {
                            members.push(walk_interface_property(m)?);
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::InterfaceDecl {
        name: name.to_string(),
        parents,
        members,
    }, span))
}

fn walk_interface_method(pair: Pair<Rule>, is_sub: bool) -> Result<InterfaceMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(InterfaceMember::Method {
        name,
        params,
        return_type,
        is_sub,
    })
}

fn walk_interface_property(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut has_read = false;
    let mut has_write = false;

    for p in pair.into_inner() {
        if p.as_rule() == Rule::property_def {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                    Rule::property_specifiers => {
                        for spec in sp.into_inner() {
                            match spec.as_rule() {
                                Rule::property_read => has_read = true,
                                Rule::property_write => has_write = true,
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(InterfaceMember::Property {
        name,
        type_hint,
        is_readonly: has_read && !has_write,
        is_writeonly: !has_read && has_write,
    })
}

// ── Enum type ──────────────────────────────────────────────────────────────

fn walk_enum_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut members = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::enum_value {
            let mut ename = String::new();
            let mut value: Option<Expression> = None;
            for ep in p.into_inner() {
                match ep.as_rule() {
                    Rule::identifier => ename = ep.as_str().to_string(),
                    Rule::expression => value = Some(walk_expression(ep)?),
                    _ => {}
                }
            }
            members.push(EnumMember { name: ename, value });
        }
    }

    Ok(Statement::with_span(StmtKind::EnumDecl {
        name: name.to_string(),
        members,
        visibility: Visibility::Public,
    }, span))
}

// ── Field declarations (in class/record bodies) ───────────────────────────

fn walk_field_decl_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            _ => {}
        }
    }

    Ok(names.into_iter().map(|n| ClassMember::Field {
        name: n,
        type_hint: type_hint.clone(),
        init: None,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
    }).collect())
}

// ── Class member signatures ────────────────────────────────────────────────

fn walk_class_constructor_sig(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut params = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                if sp.as_rule() == Rule::param_clause {
                    params = walk_param_clause(sp)?;
                }
            }
        }
    }

    Ok(ClassMember::Constructor {
        params,
        body: Vec::new(), // Body comes from method_impl
        base_args: None,
        visibility: Visibility::Public,
    })
}

fn walk_class_method_sig(pair: Pair<Rule>, is_destructor: bool) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    Rule::method_directives => {
                        walk_method_directives(sp, &mut modifiers);
                    }
                    _ => {}
                }
            }
        }
    }

    if is_destructor {
        name = "Destroy".to_string();
    }

    let is_sub = return_type.is_none();
    Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body: Vec::new(), // Body comes from method_impl
        modifiers,
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub,
    }))))
}

fn walk_class_class_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    // class procedure / class function / class var
    let _inner_text = pair.as_str().to_lowercase();
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut type_hint: Option<String> = None;
    let mut is_field = false;
    let mut modifiers = Modifiers { is_static: true, ..Default::default() };

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => name = sp.as_str().to_string(),
                        Rule::param_clause => params = walk_param_clause(sp)?,
                        Rule::return_type_clause => {
                            for rt in sp.into_inner() {
                                if rt.as_rule() == Rule::type_ref {
                                    return_type = Some(type_ref_to_string(&rt));
                                }
                            }
                        }
                        Rule::method_directives => {
                            walk_method_directives(sp, &mut modifiers);
                        }
                        _ => {}
                    }
                }
            }
            Rule::field_decl => {
                is_field = true;
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier_list => {
                            for id in sp.into_inner() {
                                if id.as_rule() == Rule::identifier {
                                    name = id.as_str().to_string();
                                }
                            }
                        }
                        Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_field {
        Ok(ClassMember::Field {
            name,
            type_hint,
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
        })
    } else {
        let is_sub = return_type.is_none();
        Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name, params, return_type, body: Vec::new(),
            modifiers, handles: Vec::new(),
            is_async: false, is_generator: false, is_sub,
        }))))
    }
}

fn walk_class_property_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut getter: Option<Vec<Statement>> = None;
    let mut setter: Option<PropertySetter> = None;
    let modifiers = Modifiers::default();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::property_def {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => {
                        if name.is_empty() {
                            name = sp.as_str().to_string();
                        }
                    }
                    Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                    Rule::property_specifiers => {
                        for spec in sp.into_inner() {
                            match spec.as_rule() {
                                Rule::property_read => {
                                    // read GetFoo → getter delegates to method
                                    let getter_name = spec.into_inner()
                                        .find(|p| p.as_rule() == Rule::identifier)
                                        .map(|p| p.as_str().to_string())
                                        .unwrap_or_default();
                                    getter = Some(vec![Statement::new(StmtKind::Return(
                                        Some(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(Expression::new(ExprKind::This)),
                                                field: getter_name,
                                                null_safe: false,
                                            })),
                                            args: Vec::new(),
                                            optional: false,
                                        }))
                                    ))]);
                                }
                                Rule::property_write => {
                                    let setter_name = spec.into_inner()
                                        .find(|p| p.as_rule() == Rule::identifier)
                                        .map(|p| p.as_str().to_string())
                                        .unwrap_or_default();
                                    let param = Param {
                                        name: "value".to_string(),
                                        type_hint: type_hint.clone(),
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    };
                                    setter = Some(PropertySetter {
                                        param,
                                        body: vec![Statement::new(StmtKind::Expr(
                                            Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::new(ExprKind::Member {
                                                    object: Box::new(Expression::new(ExprKind::This)),
                                                    field: setter_name,
                                                    null_safe: false,
                                                })),
                                                args: vec![Argument::positional(Expression::ident("value"))],
                                                optional: false,
                                            })
                                        ))],
                                    });
                                }
                                _ => {} // default, stored, nodefault
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    let is_auto = getter.is_none() && setter.is_none();
    Ok(ClassMember::Property {
        name,
        type_hint,
        getter,
        setter,
        is_auto,
        modifiers,
    })
}

fn walk_record_method_sig(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();
    let mut method_kind = "";

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_kind_keyword => method_kind = p.as_str(),
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => name = sp.as_str().to_string(),
                        Rule::param_clause => params = walk_param_clause(sp)?,
                        Rule::return_type_clause => {
                            for rt in sp.into_inner() {
                                if rt.as_rule() == Rule::type_ref {
                                    return_type = Some(type_ref_to_string(&rt));
                                }
                            }
                        }
                        Rule::method_directives => {
                            walk_method_directives(sp, &mut modifiers);
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let kind_lower = method_kind.to_lowercase();
    if kind_lower == "constructor" {
        Ok(ClassMember::Constructor {
            params,
            body: Vec::new(),
            base_args: None,
            visibility: Visibility::Public,
        })
    } else {
        let is_sub = return_type.is_none() || kind_lower == "destructor" || kind_lower == "procedure";
        Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name, params, return_type, body: Vec::new(),
            modifiers, handles: Vec::new(),
            is_async: false, is_generator: false, is_sub,
        }))))
    }
}

fn walk_record_class_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers { is_static: true, ..Default::default() };

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    Rule::method_directives => {
                        walk_method_directives(sp, &mut modifiers);
                    }
                    _ => {}
                }
            }
        }
    }

    let is_sub = return_type.is_none();
    Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name, params, return_type, body: Vec::new(),
        modifiers, handles: Vec::new(),
        is_async: false, is_generator: false, is_sub,
    }))))
}

fn walk_record_operator_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let modifiers = Modifiers { is_static: true, ..Default::default() };

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = format!("operator_{}", sp.as_str()),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name, params, return_type, body: Vec::new(),
        modifiers, handles: Vec::new(),
        is_async: false, is_generator: false, is_sub: false,
    }))))
}

fn walk_method_directives(pair: Pair<Rule>, modifiers: &mut Modifiers) {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_directive {
            let kw = p.as_str().to_lowercase();
            match kw.as_str() {
                "virtual" => modifiers.is_virtual = true,
                "override" => modifiers.is_override = true,
                "abstract" => modifiers.is_abstract = true,
                "overload" => modifiers.is_overloads = true,
                _ => {} // reintroduce, inline, cdecl, stdcall, register, dynamic
            }
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Procedure / Function declarations and method implementations
// ════════════════════════════════════════════════════════════════════════════

fn walk_procedure_decl_or_method(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_impl_proc => return walk_method_impl_proc(p, span),
            Rule::standalone_procedure => return walk_standalone_procedure(p, span),
            _ => {}
        }
    }
    Err("procedure_decl_or_method: no inner match".into())
}

fn walk_function_decl_or_method(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_impl_func => return walk_method_impl_func(p, span),
            Rule::standalone_function => return walk_standalone_function(p, span),
            _ => {}
        }
    }
    Err("function_decl_or_method: no inner match".into())
}

fn walk_method_impl_proc(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 { class_name = p.as_str().to_string(); }
                else { method_name = p.as_str().to_string(); }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name: format!("{}.{}", class_name, method_name),
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: true,
    }, span))
}

fn walk_method_impl_func(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 { class_name = p.as_str().to_string(); }
                else { method_name = p.as_str().to_string(); }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => return_type = Some(type_ref_to_string(&p)),
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name: format!("{}.{}", class_name, method_name),
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }, span))
}

fn walk_constructor_method_impl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    // constructor ClassName.Create(...)
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_impl_body_proc {
            return walk_method_impl_body_proc(p, span, true);
        }
    }
    Err("constructor_method_impl: missing body".into())
}

fn walk_destructor_method_impl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_impl_body_proc {
            return walk_method_impl_body_proc(p, span, false);
        }
    }
    Err("destructor_method_impl: missing body".into())
}

fn walk_method_impl_body_proc(pair: Pair<Rule>, span: Span, _is_constructor: bool) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 { class_name = p.as_str().to_string(); }
                else { method_name = p.as_str().to_string(); }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name: format!("{}.{}", class_name, method_name),
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: true,
    }, span))
}

fn walk_standalone_procedure(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut is_forward = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::forward_directive => is_forward = true,
            Rule::procedure_body => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::decl_section => walk_decl_section(bp, &mut body)?,
                        Rule::compound_statement => body.extend(walk_compound_statement(bp)?),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_forward {
        // Forward declarations emit an empty function
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: true,
    }, span))
}

fn walk_standalone_function(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut is_forward = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => return_type = Some(type_ref_to_string(&p)),
            Rule::forward_directive => is_forward = true,
            Rule::function_body => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::decl_section => walk_decl_section(bp, &mut body)?,
                        Rule::compound_statement => body.extend(walk_compound_statement(bp)?),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_forward {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    Ok(Statement::with_span(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }, span))
}

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_param_clause(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_list {
            return walk_param_list(p);
        }
    }
    Ok(Vec::new())
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param {
            params.extend(walk_param(p)?);
        }
    }
    Ok(params)
}

fn walk_param(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut pass_by = PassBy::Value;
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_mode => {
                let mode = p.as_str().to_lowercase();
                pass_by = match mode.as_str() {
                    "var" => PassBy::Ref,
                    "const" => PassBy::Const,
                    "out" => PassBy::Out,
                    _ => PassBy::Value,
                };
            }
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::param_default => {
                for dp in p.into_inner() {
                    if dp.as_rule() == Rule::expression {
                        default = Some(walk_expression(dp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(names.into_iter().map(|n| Param {
        name: n,
        type_hint: type_hint.clone(),
        default: default.clone(),
        pass_by,
        is_rest: false,
        is_kwargs: false,
        is_optional: default.is_some(),
        is_nullable: false,
    }).collect())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_compound_statement(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::statement_list {
            return walk_statement_list(p);
        }
    }
    Ok(Vec::new())
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::statement {
            let stmt = walk_statement(p)?;
            if !matches!(stmt.kind, StmtKind::Empty) {
                stmts.push(stmt);
            }
        }
    }
    Ok(stmts)
}

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next();
    let inner = match inner {
        Some(p) => p,
        None => return Ok(Statement::with_span(StmtKind::Empty, span)),
    };

    let kind = match inner.as_rule() {
        Rule::compound_statement => {
            StmtKind::Block(walk_compound_statement(inner)?)
        }
        Rule::if_statement => walk_if_statement(inner)?,
        Rule::for_statement => walk_for_statement(inner)?,
        Rule::for_in_statement => walk_for_in_statement(inner)?,
        Rule::while_statement => walk_while_statement(inner)?,
        Rule::repeat_statement => walk_repeat_statement(inner)?,
        Rule::case_statement => walk_case_statement(inner)?,
        Rule::with_statement => walk_with_statement(inner)?,
        Rule::try_statement => walk_try_statement(inner)?,
        Rule::raise_statement => walk_raise_statement(inner)?,
        Rule::exit_statement => walk_exit_statement(inner)?,
        Rule::halt_statement => walk_halt_statement(inner)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::inherited_statement => walk_inherited_statement(inner)?,
        Rule::assign_or_call_statement => walk_assign_or_call(inner)?,
        Rule::empty_statement => StmtKind::Empty,
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };

    Ok(Statement::with_span(kind, span))
}

// ── If ─────────────────────────────────────────────────────────────────────

fn walk_if_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    // First is always expression (condition)
    let cond = walk_expression(parts.remove(0))?;
    // Second is the then statement
    let then_stmt = walk_statement(parts.remove(0))?;
    let then_body = flatten_stmt(then_stmt);

    let else_body = if !parts.is_empty() {
        // else_clause
        let else_clause = parts.remove(0);
        let else_stmt = else_clause.into_inner().next()
            .map(|p| walk_statement(p))
            .transpose()?;
        else_stmt.map(|s| flatten_stmt(s))
    } else {
        None
    };

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── For ────────────────────────────────────────────────────────────────────

fn walk_for_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let src = pair.as_str().to_lowercase();
    let is_downto = src.contains(" downto ");

    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    // for identifier := expr (to|downto) expr do statement
    let var_name = parts.remove(0).as_str().to_string(); // identifier
    let start_expr = walk_expression(parts.remove(0))?; // start expression
    let end_expr = walk_expression(parts.remove(0))?; // end expression
    let body_stmt = walk_statement(parts.remove(0))?; // body statement

    // Build C-style for: init, cond, update
    let init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: start_expr,
    });

    let cond = Expression::new(ExprKind::Binary {
        op: if is_downto { BinOp::GtEq } else { BinOp::LtEq },
        left: Box::new(Expression::ident(&var_name)),
        right: Box::new(end_expr),
    });

    let update = Expression::new(ExprKind::Binary {
        op: if is_downto { BinOp::Sub } else { BinOp::Add },
        left: Box::new(Expression::ident(&var_name)),
        right: Box::new(Expression::int(1)),
    });
    let update_assign = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&var_name)),
        value: Box::new(update),
    });

    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update_assign),
        body: flatten_stmt(body_stmt),
    })
}

// ── For-in ─────────────────────────────────────────────────────────────────

fn walk_for_in_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let var_name = parts.remove(0).as_str().to_string();
    let iter_expr = walk_expression(parts.remove(0))?;
    let body_stmt = walk_statement(parts.remove(0))?;

    Ok(StmtKind::ForIn {
        var: var_name,
        key: None,
        iter: iter_expr,
        body: flatten_stmt(body_stmt),
        of: true, // Pascal for-in iterates values, like JS for...of
        else_body: None,
        is_async: false,
    })
}

// ── While ──────────────────────────────────────────────────────────────────

fn walk_while_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let cond = walk_expression(parts.remove(0))?;
    let body_stmt = walk_statement(parts.remove(0))?;

    Ok(StmtKind::While {
        cond,
        body: flatten_stmt(body_stmt),
        else_body: None,
    })
}

// ── Repeat/Until ───────────────────────────────────────────────────────────

fn walk_repeat_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut cond: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_list => body = walk_statement_list(p)?,
            Rule::expression => cond = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(StmtKind::DoWhile {
        body,
        cond: cond.unwrap_or_else(|| Expression::bool(true)),
        until: true,
    })
}

// ── Case ───────────────────────────────────────────────────────────────────

fn walk_case_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr: Option<Expression> = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => {
                if expr.is_none() {
                    expr = Some(walk_expression(p)?);
                }
            }
            Rule::case_arm => {
                cases.push(walk_case_arm(p)?);
            }
            Rule::case_else => {
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::statement_list {
                        default = Some(walk_statement_list(cp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: expr.unwrap_or_else(|| Expression::null()),
        cases,
        default,
    })
}

fn walk_case_arm(pair: Pair<Rule>) -> Result<SwitchCase, String> {
    let mut conditions = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::case_value_list => {
                for cv in p.into_inner() {
                    if cv.as_rule() == Rule::case_value {
                        conditions.push(walk_case_value(cv)?);
                    }
                }
            }
            Rule::case_arm_body => {
                let inner = cv_first(p)?;
                match inner.as_rule() {
                    Rule::compound_statement => body = walk_compound_statement(inner)?,
                    Rule::statement => body = flatten_stmt(walk_statement(inner)?),
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(SwitchCase { conditions, body })
}

fn walk_case_value(pair: Pair<Rule>) -> Result<CaseCondition, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    if parts.len() == 2 {
        // Range: expr..expr
        let from = walk_expression(parts[0].clone())?;
        let to = walk_expression(parts[1].clone())?;
        Ok(CaseCondition::Range { from, to })
    } else if parts.len() == 1 {
        let val = walk_expression(parts[0].clone())?;
        Ok(CaseCondition::Value(val))
    } else {
        Err("Empty case_value".into())
    }
}

// ── With ───────────────────────────────────────────────────────────────────

fn walk_with_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut items = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => {
                items.push(WithItem {
                    expr: walk_expression(p)?,
                    var: None,
                });
            }
            Rule::statement => {
                body = flatten_stmt(walk_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::With {
        items,
        body,
        is_async: false,
    })
}

// ── Try ────────────────────────────────────────────────────────────────────

fn walk_try_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_list => {
                body = walk_statement_list(p)?;
            }
            Rule::try_handler => {
                for hp in p.into_inner() {
                    match hp.as_rule() {
                        Rule::except_handler => {
                            for ep in hp.into_inner() {
                                if ep.as_rule() == Rule::except_body {
                                    catches = walk_except_body(ep)?;
                                }
                            }
                        }
                        Rule::finally_handler => {
                            for fp in hp.into_inner() {
                                if fp.as_rule() == Rule::statement_list {
                                    finally = Some(walk_statement_list(fp)?);
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

    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

fn walk_except_body(pair: Pair<Rule>) -> Result<Vec<CatchClause>, String> {
    let mut clauses = Vec::new();
    let mut has_on_clauses = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::on_clause => {
                has_on_clauses = true;
                clauses.push(walk_on_clause(p)?);
            }
            Rule::except_else => {
                // else clause in except → catch-all
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::statement_list {
                        clauses.push(CatchClause {
                            types: Vec::new(),
                            var_name: None,
                            stack_var: None,
                            body: walk_statement_list(ep)?,
                            when_clause: None,
                        });
                    }
                }
            }
            Rule::statement_list => {
                if !has_on_clauses {
                    // Bare except with just a statement list → catch-all
                    clauses.push(CatchClause {
                        types: Vec::new(),
                        var_name: None,
                        stack_var: None,
                        body: walk_statement_list(p)?,
                        when_clause: None,
                    });
                }
            }
            _ => {}
        }
    }

    Ok(clauses)
}

fn walk_on_clause(pair: Pair<Rule>) -> Result<CatchClause, String> {
    let mut var_name: Option<String> = None;
    let mut type_name = String::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::on_var_binding => {
                for vp in p.into_inner() {
                    if vp.as_rule() == Rule::identifier {
                        var_name = Some(vp.as_str().to_string());
                    }
                }
            }
            Rule::identifier => type_name = p.as_str().to_string(),
            Rule::statement => body = flatten_stmt(walk_statement(p)?),
            _ => {}
        }
    }

    Ok(CatchClause {
        types: vec![type_name],
        var_name,
        stack_var: None,
        body,
        when_clause: None,
    })
}

// ── Raise ──────────────────────────────────────────────────────────────────

fn walk_raise_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    Ok(StmtKind::Throw { expr, cause: None })
}

// ── Exit ───────────────────────────────────────────────────────────────────

fn walk_exit_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    Ok(StmtKind::Return(expr))
}

// ── Halt ───────────────────────────────────────────────────────────────────

fn walk_halt_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    // Halt maps to a special call; emit as Expr(Call(Halt, [code]))
    let args = expr.into_iter().map(Argument::positional).collect();
    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Halt")),
        args,
        optional: false,
    })))
}

// ── Inherited statement ────────────────────────────────────────────────────

fn walk_inherited_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut method: Option<String> = None;
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => method = Some(p.as_str().to_string()),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::Expr(Expression::new(ExprKind::SuperCall {
        method,
        args,
    })))
}

// ── Assign or call ─────────────────────────────────────────────────────────

fn walk_assign_or_call(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let src = pair.as_str();
    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();

    if parts.len() == 1 {
        // Pure expression used as statement (procedure call, etc.)
        let expr = walk_expression(parts.into_iter().next().unwrap())?;

        // Pascal `FreeAndNil(x)` is sugar for `x := nil` — we have GC, so the
        // free is a no-op but the variable still needs to be cleared so that
        // `Assigned(x)` returns false afterwards. Rewrite at the walker.
        if let ExprKind::Call { callee, args, .. } = &expr.kind {
            if let ExprKind::Ident(name) = &callee.kind {
                if name.eq_ignore_ascii_case("FreeAndNil") && args.len() == 1 {
                    return Ok(StmtKind::Assign {
                        targets: vec![args[0].value.clone()],
                        value: Expression::null(),
                    });
                }
            }
        }

        // Pascal allows zero-arg procedure calls without parens: `Hello;` means
        // `Hello();`. At statement level, a bare identifier or member access
        // that isn't already a Call is implicitly a zero-arg invocation.
        let expr = match expr.kind {
            ExprKind::Call { .. }
            | ExprKind::New { .. }
            | ExprKind::Assign { .. }
            | ExprKind::Lit(_) => expr,
            ExprKind::Ident(_) | ExprKind::Member { .. } => Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(expr.clone()),
                    args: Vec::new(),
                    optional: false,
                },
                expr.span,
            ),
            _ => expr,
        };
        return Ok(StmtKind::Expr(expr));
    }

    if parts.len() >= 2 {
        let target_pair = parts[0].clone();
        let target = walk_expression(target_pair)?;

        // Check for compound assignment operators
        // The grammar captures expression ~ (":=" | "+=" | "-=" | "*=" | "/=") ~ expression
        // After the first expression, remaining pairs are: potentially just the value expression
        // But the operator is part of the rule text, not a separate pair.
        // We need to detect the operator from the source text.

        let value_pair = parts.last().unwrap().clone();
        let value = walk_expression(value_pair)?;

        if src.contains(":=") {
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value,
            });
        } else if src.contains("+=") {
            return Ok(StmtKind::CompoundAssign {
                target, op: CompoundOp::Add, value,
            });
        } else if src.contains("-=") {
            return Ok(StmtKind::CompoundAssign {
                target, op: CompoundOp::Sub, value,
            });
        } else if src.contains("*=") {
            return Ok(StmtKind::CompoundAssign {
                target, op: CompoundOp::Mul, value,
            });
        } else if src.contains("/=") {
            return Ok(StmtKind::CompoundAssign {
                target, op: CompoundOp::Div, value,
            });
        }

        // Fallback: assignment
        return Ok(StmtKind::Assign {
            targets: vec![target],
            value,
        });
    }

    Ok(StmtKind::Empty)
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::expression => {
            // expression = { is_as_expression }
            let inner = pair.into_inner().next().ok_or("Empty expression")?;
            walk_expr_kind(inner)
        }

        Rule::is_as_expression => {
            // is_as_expression = { relational ~ (is_as_op ~ identifier)* }
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }

            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i + 1 < inner.len() {
                let op_pair = &inner[i];
                let op_str = op_pair.as_str().to_lowercase();
                let type_name = inner[i + 1].as_str().to_string();

                if op_str == "is" {
                    left = Expression::new(ExprKind::IsType {
                        expr: Box::new(left),
                        type_name,
                    });
                } else {
                    // as
                    left = Expression::new(ExprKind::Cast {
                        expr: Box::new(left),
                        type_name,
                    });
                }
                i += 2;
            }
            Ok(left.kind)
        }

        Rule::relational => {
            walk_binary_chain(pair, |op_str| {
                match op_str {
                    "<>" => BinOp::NotEq,
                    "<=" => BinOp::LtEq,
                    ">=" => BinOp::GtEq,
                    "<" => BinOp::Lt,
                    ">" => BinOp::Gt,
                    "=" => BinOp::Eq,
                    s if s.starts_with("in") => BinOp::In,
                    _ => BinOp::Eq,
                }
            })
        }

        Rule::additive => {
            walk_binary_chain(pair, |op_str| {
                match op_str {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    s if s.starts_with("or") => BinOp::Or,
                    s if s.starts_with("xor") => BinOp::BitXor,
                    _ => BinOp::Add,
                }
            })
        }

        Rule::multiplicative => {
            walk_binary_chain(pair, |op_str| {
                match op_str {
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    s if s.starts_with("div") => BinOp::IDiv,
                    s if s.starts_with("mod") => BinOp::Mod,
                    s if s.starts_with("and") => BinOp::And,
                    s if s.starts_with("shl") => BinOp::Shl,
                    s if s.starts_with("shr") => BinOp::Shr,
                    _ => BinOp::Mul,
                }
            })
        }

        Rule::unary => {
            // Pest does not include literal token matches (like "-", "@") as inner
            // pairs — they're consumed silently. Inspect the source text to decide
            // whether this unary node carries a prefix operator.
            let src = pair.as_str().trim_start();
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            // Always exactly one inner pair: either the inner `unary` (when there's
            // a prefix) or the `postfix` (no prefix).
            let operand_pair = inner.pop().ok_or("Empty unary")?;
            let operand = walk_expression(operand_pair)?;

            if src.starts_with('-') {
                Ok(ExprKind::Unary { op: UnaryOp::Neg, expr: Box::new(operand) })
            } else if src.len() >= 3 && src[..3].eq_ignore_ascii_case("not")
                && !src.chars().nth(3).map_or(false, |c| c.is_alphanumeric() || c == '_')
            {
                Ok(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(operand) })
            } else if src.starts_with('@') {
                Ok(ExprKind::Unary { op: UnaryOp::AddrOf, expr: Box::new(operand) })
            } else {
                Ok(operand.kind)
            }
        }

        Rule::postfix => {
            // postfix = { primary ~ postfix_op* }
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                return Err("Empty postfix".into());
            }

            let mut expr = walk_expression(inner.remove(0))?;

            for op in inner {
                if op.as_rule() != Rule::postfix_op {
                    continue;
                }
                expr = walk_postfix_op(expr, op)?;
            }

            Ok(expr.kind)
        }

        Rule::primary => walk_primary(pair),

        // Passthrough for operator pairs that appear in binary chains
        Rule::relational_op | Rule::additive_op | Rule::multiplicative_op | Rule::is_as_op => {
            Err(format!("Operator {:?} should not be walked as expression", pair.as_rule()))
        }

        // Literals and identifiers that might appear directly
        Rule::int_literal => {
            let s = pair.as_str();
            if s.starts_with('$') {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[1..], 16).unwrap_or(0)
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }
        Rule::real_literal => {
            Ok(ExprKind::Lit(Literal::Float(pair.as_str().parse().unwrap_or(0.0))))
        }
        Rule::string_literal => {
            let raw = pair.as_str();
            // Strip surrounding quotes and unescape ''
            let inner = &raw[1..raw.len()-1];
            Ok(ExprKind::Lit(Literal::Str(inner.replace("''", "'").to_string())))
        }
        Rule::char_literal => {
            // #65 → 'A'
            let s = pair.as_str();
            let code: u32 = s[1..].parse().unwrap_or(0);
            Ok(ExprKind::Lit(Literal::Char(char::from_u32(code).unwrap_or('\0'))))
        }
        Rule::identifier => {
            Ok(ExprKind::Ident(pair.as_str().to_string()))
        }

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // Primary can be a keyword literal or have inner pairs
    let src = pair.as_str().trim();
    let src_lower = src.to_lowercase();

    // Check for keyword literals that pest may not produce inner pairs for
    match src_lower.as_str() {
        "true" => return Ok(ExprKind::Lit(Literal::Bool(true))),
        "false" => return Ok(ExprKind::Lit(Literal::Bool(false))),
        "nil" => return Ok(ExprKind::Lit(Literal::Null)),
        "result" => return Ok(ExprKind::Ident("Result".to_string())),
        _ => {}
    }

    let inner = pair.into_inner().next();
    let inner = match inner {
        Some(p) => p,
        None => {
            // If no inner pair, the whole primary text is an identifier or keyword
            // (pest sometimes doesn't create inner pairs for case-insensitive keyword matches)
            return Ok(ExprKind::Ident(src.to_string()));
        }
    };

    match inner.as_rule() {
        Rule::int_literal => walk_expr_kind(inner),
        Rule::real_literal => walk_expr_kind(inner),
        Rule::string_literal => walk_expr_kind(inner),
        Rule::char_literal => walk_expr_kind(inner),
        Rule::identifier => Ok(ExprKind::Ident(inner.as_str().to_string())),
        Rule::set_literal => walk_set_literal(inner),
        Rule::new_expression => walk_new_expression(inner),
        Rule::lambda_procedure => walk_lambda_procedure(inner),
        Rule::lambda_function => walk_lambda_function(inner),
        Rule::inherited_expression => walk_inherited_expression(inner),
        Rule::type_cast_builtin => walk_type_cast_builtin(inner),
        Rule::true_keyword => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_keyword => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::nil_keyword => Ok(ExprKind::Lit(Literal::Null)),
        Rule::result_keyword => Ok(ExprKind::Ident("Result".to_string())),
        Rule::expression => {
            // Parenthesized expression: "(" ~ expression ~ ")"
            walk_expr_kind(inner)
        }
        other => Err(format!("Unexpected primary inner: {:?}", other)),
    }
}

// ── Postfix operations ─────────────────────────────────────────────────────

fn walk_postfix_op(expr: Expression, op: Pair<Rule>) -> Result<Expression, String> {
    let op_src = op.as_str();
    let parts: Vec<Pair<Rule>> = op.into_inner().collect();

    if op_src == "^" {
        // Dereference: ptr^
        return Ok(Expression::new(ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: Box::new(expr),
        }));
    }

    if op_src.starts_with('.') {
        // Field access or method call: obj.Field or obj.Method(args)
        // Grammar: "." ~ identifier ~ arg_list  |  "." ~ identifier
        let mut ident = String::new();
        let mut arg_list: Option<Pair<Rule>> = None;

        for p in &parts {
            match p.as_rule() {
                Rule::identifier => ident = p.as_str().to_string(),
                Rule::arg_list => arg_list = Some(p.clone()),
                _ => {}
            }
        }

        // Canonicalize property-style access for builtins (e.g. arr.Length →
        // __len__(arr)) so the compiler dispatches via compiler_common::canonical.
        // Only when there are no parens — `obj.Length(...)` is a real method call.
        if arg_list.is_none() {
            if let Some(canonical) = canonicalize_pascal_member(&ident) {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(canonical)),
                    args: vec![Argument::positional(expr)],
                    optional: false,
                }));
            }
        }

        let member = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: ident,
            null_safe: false,
        });

        if let Some(al) = arg_list {
            let args = walk_arg_list(al)?;
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args,
                optional: false,
            }));
        } else {
            return Ok(member);
        }
    }

    if op_src.starts_with('[') {
        // Index access: arr[i]
        let index_expr = parts.into_iter()
            .find(|p| p.as_rule() == Rule::expression)
            .map(walk_expression)
            .transpose()?
            .unwrap_or_else(|| Expression::int(0));
        return Ok(Expression::new(ExprKind::Index {
            object: Box::new(expr),
            index: Box::new(index_expr),
        }));
    }

    if op_src.starts_with('(') {
        // Function call: F(args)
        let args = parts.into_iter()
            .find(|p| p.as_rule() == Rule::arg_list)
            .map(walk_arg_list)
            .transpose()?
            .unwrap_or_default();

        // Canonicalize Pascal's function-style builtins to canonical names so the
        // compiler can dispatch them via compiler_common::canonical regardless of
        // source language. Pascal is case-insensitive.
        if let ExprKind::Ident(name) = &expr.kind {
            if let Some(canonical) = canonicalize_pascal_builtin(name, args.len()) {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(canonical)),
                    args,
                    optional: false,
                }));
            }
        }

        // Check if the callee is an identifier that looks like a type cast
        // (e.g. Integer(x), String(x)) — the grammar handles builtin type casts via
        // type_cast_builtin, but identifier-based type casts (e.g. TMyType(x))
        // are just calls — let the compiler handle semantics.
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(expr),
            args,
            optional: false,
        }));
    }

    // Fallback
    Ok(expr)
}

// ── Canonical builtin normalization ────────────────────────────────────────

/// Normalize Pascal's function-style builtins to canonical names so the compiler
/// can dispatch them through `compiler_common::canonical`. This keeps language
/// surface syntax in the walker; the compiler stays language-agnostic.
fn canonicalize_pascal_builtin(name: &str, argc: usize) -> Option<&'static str> {
    match (name.to_lowercase().as_str(), argc) {
        ("length", 1) => Some("__len__"),
        _ => None,
    }
}

/// Pascal property-style member access canonicalization (case-insensitive).
fn canonicalize_pascal_member(name: &str) -> Option<&'static str> {
    match name.to_lowercase().as_str() {
        "length" | "count" => Some("__len__"),
        _ => None,
    }
}

// ── Argument list ──────────────────────────────────────────────────────────

fn walk_arg_list(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            let value = walk_expression(p)?;
            Ok(Argument::positional(value))
        })
        .collect()
}

// ── Set literal ────────────────────────────────────────────────────────────

fn walk_set_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements: Vec<ArrayElement> = pair.into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            let value = walk_expression(p)?;
            Ok(ArrayElement { key: None, value, spread: false, by_ref: false })
        })
        .collect::<Result<_, String>>()?;
    Ok(ExprKind::Array(elements))
}

// ── New expression ─────────────────────────────────────────────────────────

fn walk_new_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut class_name = String::new();
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => class_name = p.as_str().to_string(),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    Ok(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args,
    })
}

// ── Lambda expressions ─────────────────────────────────────────────────────

fn walk_lambda_procedure(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body = LambdaBody::Block(Vec::new());

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::compound_statement => {
                body = LambdaBody::Block(walk_compound_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: Vec::new(),
    })
}

fn walk_lambda_function(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body = LambdaBody::Block(Vec::new());

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => { /* return type hint — ignored for lambda body */ }
            Rule::compound_statement => {
                body = LambdaBody::Block(walk_compound_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: Vec::new(),
    })
}

// ── Inherited expression ───────────────────────────────────────────────────

fn walk_inherited_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut method: Option<String> = None;
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => method = Some(p.as_str().to_string()),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    if method.is_some() || !args.is_empty() {
        Ok(ExprKind::SuperCall { method, args })
    } else {
        // Bare `inherited` → Super
        Ok(ExprKind::Super)
    }
}

// ── Type cast with builtin type ────────────────────────────────────────────

fn walk_type_cast_builtin(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut type_name = String::new();
    let mut expr: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::builtin_type => type_name = p.as_str().to_string(),
            Rule::expression => expr = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::null())),
        type_name,
    })
}

// ── Binary chain helper ────────────────────────────────────────────────────

fn walk_binary_chain<F>(pair: Pair<Rule>, map_op: F) -> Result<ExprKind, String>
where
    F: Fn(&str) -> BinOp,
{
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    // First operand
    let mut left = walk_expression(inner.remove(0))?;

    // Remaining: (op, operand) pairs
    let mut i = 0;
    while i + 1 < inner.len() {
        let op_str = inner[i].as_str().trim().to_lowercase();
        let right = walk_expression(inner[i + 1].clone())?;
        let bin_op = map_op(&op_str);

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        i += 2;
    }

    Ok(left.kind)
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

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

fn type_ref_to_string(pair: &Pair<Rule>) -> String {
    pair.as_str().trim().to_string()
}

/// Flatten a single statement into a Vec — if it's a Block, unwrap it.
fn flatten_stmt(stmt: Statement) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::Block(stmts) => stmts,
        _ => vec![stmt],
    }
}

/// Get first inner pair from a compound pair.
fn cv_first(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    pair.into_inner().next().ok_or_else(|| "Expected inner pair".to_string())
}
