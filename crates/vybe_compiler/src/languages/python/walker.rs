use super::{PythonParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

// ════════════════════════════════════════════════════════════════════════════
// Indentation preprocessor
// ════════════════════════════════════════════════════════════════════════════
// Python uses indentation for blocks. pest cannot track indent state, so we
// insert explicit markers before parsing:
//   ⇥ (U+21E5) = INDENT
//   ⇤ (U+21E4) = DEDENT

fn preprocess_indentation(source: &str) -> String {
    // Phase 1: Resolve physical lines into logical lines.
    // Handles explicit continuation (backslash) and implicit continuation (unclosed brackets).
    let mut logical_lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut bracket_depth: i32 = 0;

    for line in source.lines() {
        let trimmed = line.trim();

        // During continuation, skip blank/comment lines
        if !current.is_empty() && (trimmed.is_empty() || trimmed.starts_with('#')) {
            continue;
        }

        if current.is_empty() {
            current.push_str(line);
        } else {
            // Continuation: join with space + trimmed content
            current.push(' ');
            current.push_str(trimmed);
        }

        // Update bracket depth (skip chars after # comment marker)
        for c in line.chars() {
            match c {
                '(' | '[' | '{' => bracket_depth += 1,
                ')' | ']' | '}' => bracket_depth -= 1,
                '#' => break,
                _ => {}
            }
        }

        // Explicit continuation: backslash at end of line
        if line.trim_end().ends_with('\\') {
            if let Some(pos) = current.rfind('\\') {
                current.truncate(pos);
            }
            continue;
        }

        // Implicit continuation: unclosed brackets
        if bracket_depth > 0 {
            continue;
        }

        logical_lines.push(std::mem::take(&mut current));
        bracket_depth = 0;
    }
    if !current.is_empty() {
        logical_lines.push(current);
    }

    // Phase 2: Process indentation on logical lines
    let mut result = String::with_capacity(source.len() * 2);
    let mut indent_stack: Vec<usize> = vec![0];
    let mut first = true;

    for line in &logical_lines {
        // Count leading spaces (expand tabs to 8)
        let mut indent = 0;
        let mut chars = line.chars().peekable();
        while let Some(&c) = chars.peek() {
            match c {
                ' ' => {
                    indent += 1;
                    chars.next();
                }
                '\t' => {
                    indent += 8 - (indent % 8);
                    chars.next();
                }
                _ => break,
            }
        }

        let rest: String = chars.collect();

        // Skip blank lines and comment-only lines
        let trimmed = rest.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            if !first {
                result.push('\n');
            }
            result.push_str(line);
            first = false;
            continue;
        }

        if !first {
            result.push('\n');
        }
        first = false;

        let current_indent = *indent_stack.last().unwrap();

        if indent > current_indent {
            indent_stack.push(indent);
            result.push('\u{21E5}'); // INDENT
        } else {
            while indent < *indent_stack.last().unwrap() {
                indent_stack.pop();
                result.push('\u{21E4}'); // DEDENT
            }
            // After popping, if indent is above the new top, it's a new block level
            if indent > *indent_stack.last().unwrap() {
                indent_stack.push(indent);
                result.push('\u{21E5}'); // INDENT
            }
        }

        result.push_str(line);
    }

    // Close remaining indents at EOF
    while indent_stack.len() > 1 {
        indent_stack.pop();
        result.push('\n');
        result.push('\u{21E4}');
    }

    result
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let preprocessed = preprocess_indentation(source);
    let pairs = PythonParser::parse(Rule::program, &preprocessed)
        .map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                walk_stmt_into(top, &mut body, &mut imports)?;
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                _ => walk_stmt_into(pair, &mut body, &mut imports)?,
            }
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Python,
        body,
        imports,
    })
}

fn walk_stmt_into(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    imports: &mut Vec<Import>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::import_stmt => imports.push(walk_import(pair)?),
        Rule::import_from_stmt => imports.push(walk_import_from(pair)?),
        _ => body.push(walk_statement(pair)?),
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::pass_stmt => StmtKind::Empty,
        Rule::break_stmt | Rule::break_inline => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_stmt | Rule::continue_inline => StmtKind::Continue(ContinueTarget::Implicit),

        Rule::function_def => walk_func_def(pair, false, Vec::new())?,
        Rule::class_def => walk_class_def(pair, Vec::new())?,
        Rule::decorated_def => walk_decorated(pair)?,
        Rule::async_stmt => walk_async_stmt(pair)?,

        Rule::if_stmt => walk_if(pair)?,
        Rule::while_stmt => walk_while(pair)?,
        Rule::for_stmt => walk_for(pair, false)?,
        Rule::try_stmt => walk_try(pair)?,
        Rule::with_stmt => walk_with(pair, false)?,
        Rule::match_stmt => walk_match(pair)?,

        Rule::return_stmt | Rule::return_inline => walk_return(pair)?,
        Rule::raise_stmt | Rule::raise_inline => walk_raise(pair)?,
        Rule::del_stmt => walk_del(pair)?,
        Rule::assert_stmt => walk_assert(pair)?,
        Rule::global_stmt => walk_scope_decl(pair, ScopeDeclKind::Global)?,
        Rule::nonlocal_stmt => walk_scope_decl(pair, ScopeDeclKind::Nonlocal)?,

        Rule::import_stmt => return Ok(Statement::new(StmtKind::Empty)), // handled in walk_stmt_into
        Rule::import_from_stmt => return Ok(Statement::new(StmtKind::Empty)),

        Rule::expr_or_assign_stmt | Rule::expr_or_assign_inline => walk_expr_or_assign(pair)?,

        Rule::pass_inline => StmtKind::Empty,

        Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => StmtKind::Empty,

        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

// ── Generator → eager collection helpers ────────────────────────────────────

/// Recursively check if a statement list contains any Yield expressions.
fn body_has_yield(stmts: &[Statement]) -> bool {
    stmts.iter().any(|s| stmt_has_yield(s))
}

fn stmt_has_yield(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_has_yield(e),
        StmtKind::Return(Some(e)) => expr_has_yield(e),
        StmtKind::Assign { value, .. } => expr_has_yield(value),
        StmtKind::If {
            cond,
            then_body,
            else_body,
            ..
        } => {
            expr_has_yield(cond)
                || body_has_yield(then_body)
                || else_body.as_ref().map_or(false, |eb| body_has_yield(eb))
        }
        StmtKind::While { cond, body, .. } => expr_has_yield(cond) || body_has_yield(body),
        StmtKind::ForIn { body, .. } => body_has_yield(body),
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            body_has_yield(body)
                || catches.iter().any(|cb| body_has_yield(&cb.body))
                || finally.as_ref().map_or(false, |fb| body_has_yield(fb))
        }
        StmtKind::With { body, .. } => body_has_yield(body),
        _ => false,
    }
}

fn expr_has_yield(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
        ExprKind::Call { args, .. } => args.iter().any(|a| expr_has_yield(&a.value)),
        ExprKind::Binary { left, right, .. } => expr_has_yield(left) || expr_has_yield(right),
        ExprKind::Unary { expr: e, .. } => expr_has_yield(e),
        ExprKind::Index { object, index, .. } => expr_has_yield(object) || expr_has_yield(index),
        _ => false,
    }
}

/// Rewrite a generator body: prepend `__gen_result = []`, replace `yield X` with
/// `__gen_result.push(X)`, append `return __gen_result`.
fn rewrite_generator_body(stmts: Vec<Statement>) -> Vec<Statement> {
    let gen_var = "__gen_result";
    let mut out = Vec::new();

    // __gen_result = []
    out.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(gen_var)],
        value: Expression::new(ExprKind::Array(Vec::new())),
    }));

    // Transform body
    for stmt in stmts {
        out.extend(rewrite_stmt_yields(stmt, gen_var));
    }

    // return __gen_result
    out.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        gen_var,
    )))));

    out
}

fn rewrite_stmt_yields(stmt: Statement, gen_var: &str) -> Vec<Statement> {
    match stmt.kind {
        // Bare `yield X` as expression statement → __gen_result.push(X)
        StmtKind::Expr(ref e) => {
            if let ExprKind::Yield(Some(val)) = &e.kind {
                return vec![make_push(gen_var, *val.clone())];
            }
            if let ExprKind::Yield(None) = &e.kind {
                return vec![make_push(gen_var, Expression::null())];
            }
            vec![stmt]
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            vec![Statement::new(StmtKind::If {
                cond,
                then_body: then_body
                    .into_iter()
                    .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                    .collect(),
                elifs: elifs
                    .into_iter()
                    .map(|(c, b)| {
                        (
                            c,
                            b.into_iter()
                                .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                                .collect(),
                        )
                    })
                    .collect(),
                else_body: else_body.map(|eb| {
                    eb.into_iter()
                        .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                        .collect()
                }),
            })]
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            vec![Statement::new(StmtKind::While {
                cond,
                body: body
                    .into_iter()
                    .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                    .collect(),
                else_body,
            })]
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            vec![Statement::new(StmtKind::ForIn {
                var,
                key,
                iter,
                of,
                is_async,
                body: body
                    .into_iter()
                    .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                    .collect(),
                else_body: else_body.map(|eb| {
                    eb.into_iter()
                        .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                        .collect()
                }),
            })]
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            vec![Statement::new(StmtKind::Try {
                body: body
                    .into_iter()
                    .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                    .collect(),
                catches: catches
                    .into_iter()
                    .map(|cb| CatchClause {
                        body: cb
                            .body
                            .into_iter()
                            .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                            .collect(),
                        ..cb
                    })
                    .collect(),
                else_body,
                finally: finally.map(|fb| {
                    fb.into_iter()
                        .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                        .collect()
                }),
            })]
        }
        StmtKind::With {
            items,
            body,
            is_async,
        } => {
            vec![Statement::new(StmtKind::With {
                items,
                is_async,
                body: body
                    .into_iter()
                    .flat_map(|s| rewrite_stmt_yields(s, gen_var))
                    .collect(),
            })]
        }
        _ => vec![stmt],
    }
}

fn make_push(gen_var: &str, val: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(gen_var)),
            field: "append".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(val)],
        optional: false,
    })))
}

// ── Function def ────────────────────────────────────────────────────────────

fn walk_func_def(
    pair: Pair<Rule>,
    is_async: bool,
    decorators: Vec<Expression>,
) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::block => body = walk_block(p)?,
            Rule::expression
            | Rule::named_expr
            | Rule::ternary_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::power
            | Rule::await_expr
            | Rule::postfix
            | Rule::primary => {
                // return type annotation — just note it as string
                return_type = Some(p.as_str().to_string());
            }
            _ => {}
        }
    }

    // Generator function: transform yield statements into eager collection.
    // def gen(): yield 1; yield 2 → def gen(): __gen_result = []; __gen_result.append(1); ...; return __gen_result
    // Generators: two lowering paths, chosen by the function's
    // decorator list.
    //   * Default — eager-list rewrite: yields append to a list that
    //     is returned at the end, so `for v in gen()` iterates the
    //     list via the standard for-in protocol. Backwards-compatible
    //     with the existing generator test suite.
    //   * `@generator` decorator — true lazy generator via the
    //     stack-switching proposal: the function compiles with
    //     `is_generator = true`, calls return a `Continuation`, and
    //     each `yield` compiles to a `SUSPEND` opcode. Consuming
    //     requires explicit `RESUME` (or a future iterator-protocol-
    //     aware for-in) — no automatic eager materialisation.
    let has_yield = body_has_yield(&body);
    let wants_true_generator = decorators.iter().any(|d| match &d.kind {
        ExprKind::Ident(n) => n.eq_ignore_ascii_case("generator"),
        _ => false,
    });
    if has_yield && !wants_true_generator {
        body = rewrite_generator_body(body);
    }

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers {
            decorators,
            ..Default::default()
        },
        handles: Vec::new(),
        is_async,
        is_generator: has_yield && wants_true_generator,
        is_sub: false,
    })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_item {
            let inner = p.into_inner().next();
            if let Some(item) = inner {
                match item.as_rule() {
                    Rule::normal_param => {
                        let mut name = String::new();
                        let mut default = None;
                        let mut type_hint = None;
                        let mut seen_first_expr = false;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ => {
                                    if !seen_first_expr {
                                        // First expression after identifier could be type annotation
                                        type_hint = Some(c.as_str().to_string());
                                        seen_first_expr = true;
                                    } else {
                                        default = Some(walk_expression(c)?);
                                    }
                                }
                            }
                        }
                        // If we only saw one expr and there's no "=", it was actually the type hint
                        // We need to check if there was a default — if type_hint is set but default is not,
                        // check if the param text had "="
                        params.push(Param {
                            name,
                            type_hint,
                            is_optional: default.is_some(),
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_nullable: false,
                        });
                    }
                    Rule::star_param => {
                        let mut name = String::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier {
                                name = c.as_str().to_string();
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: true,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::double_star_param => {
                        let mut name = String::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier {
                                name = c.as_str().to_string();
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: true,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::bare_star | Rule::slash_param => {} // separator, not a param
                    _ => {}
                }
            }
        }
    }
    Ok(params)
}

// ── Class def ───────────────────────────────────────────────────────────────

fn walk_class_def(pair: Pair<Rule>, _decorators: Vec<Expression>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut body_stmts = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::class_arg_list => {
                for arg in p.into_inner() {
                    if arg.as_rule() == Rule::class_arg {
                        // Just extract as string base name
                        let text = arg.as_str().trim().to_string();
                        if !text.contains('=') && !text.starts_with("**") {
                            parents.push(text);
                        }
                    }
                }
            }
            Rule::block => body_stmts = walk_block(p)?,
            _ => {}
        }
    }

    // Convert body statements into ClassMembers
    let members = stmts_to_class_members(body_stmts);

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn stmts_to_class_members(stmts: Vec<Statement>) -> Vec<ClassMember> {
    let mut members: Vec<ClassMember> = Vec::new();
    // Track Property member index by name so @x.setter can find the getter.
    let mut property_indices: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();

    for stmt in stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                body,
                modifiers,
                is_async,
                ..
            } => {
                if name == "__init__" {
                    // Constructor — keep `self` param; compiler strips it
                    // via `NormalClass.explicit_self_param` (set by
                    // normalize_class).
                    members.push(ClassMember::Constructor {
                        params: params.clone(),
                        body: body.clone(),
                        base_args: None,
                        initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                        visibility: Visibility::Public,
                    });
                    continue;
                }

                // Check for @property decorator → build Property getter
                let has_property = modifiers
                    .decorators
                    .iter()
                    .any(|d| matches!(&d.kind, ExprKind::Ident(n) if n == "property"));
                if has_property {
                    let idx = members.len();
                    members.push(ClassMember::Property {
                        name: name.clone(),
                        type_hint: None,
                        getter: Some(body.clone()),
                        setter: None,
                        is_auto: false,
                        modifiers: Modifiers::default(),
                    });
                    property_indices.insert(name.clone(), idx);
                    continue;
                }

                // Check for @x.setter or @x.deleter → add to existing Property
                let setter_target = modifiers.decorators.iter().find_map(|d| {
                    if let ExprKind::Member { object, field, .. } = &d.kind {
                        if field == "setter" {
                            if let ExprKind::Ident(prop_name) = &object.kind {
                                return Some((prop_name.clone(), "setter"));
                            }
                        }
                    }
                    None
                });
                if let Some((prop_name, "setter")) = setter_target {
                    if let Some(&prop_idx) = property_indices.get(&prop_name) {
                        if let ClassMember::Property { setter, .. } = &mut members[prop_idx] {
                            // Second param (after self) is the value param
                            let value_param = params.iter().nth(1).cloned().unwrap_or(Param {
                                name: "value".to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                            *setter = Some(crate::ast::PropertySetter {
                                param: value_param,
                                body: body.clone(),
                            });
                        }
                    }
                    continue;
                }

                // Method — keep `self`/`cls` param; compiler strips via
                // `NormalClass.explicit_self_param`.
                let has_staticmethod = modifiers
                    .decorators
                    .iter()
                    .any(|d| matches!(&d.kind, ExprKind::Ident(n) if n == "staticmethod"));
                // For @staticmethod, prepend a dummy "self" so that
                // explicit_self_param's skip(1) removes the dummy, keeping
                // the real params intact. Without this, skip(1) would drop
                // the first real param (e.g. `a` in `def add(a, b)`).
                let final_params = if has_staticmethod {
                    let dummy = Param {
                        name: "self".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    };
                    let mut p = vec![dummy];
                    p.extend_from_slice(params);
                    p
                } else {
                    params.clone()
                };
                let is_static =
                    has_staticmethod || final_params.first().map_or(true, |p| p.name != "self");
                let mut mods = modifiers.clone();
                mods.is_static = is_static;
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: name.clone(),
                        params: final_params,
                        return_type: None,
                        body: body.clone(),
                        modifiers: mods,
                        handles: Vec::new(),
                        is_async: *is_async,
                        is_generator: false,
                        is_sub: false,
                    },
                ))));
            }
            StmtKind::Assign { targets, value } => {
                // Class-level assignment → static Field (Python class variables)
                for target in targets {
                    if let ExprKind::Ident(field_name) = &target.kind {
                        let mut mods = Modifiers::default();
                        mods.is_static = true; // Python class-level vars are class attributes
                        members.push(ClassMember::Field {
                            name: field_name.clone(),
                            type_hint: None,
                            init: Some(value.clone()),
                            modifiers: mods,
                            with_events: false,
                            array_bounds: None,
                        });
                    }
                }
            }
            StmtKind::Empty => {} // pass
            _ => {
                // Nested class or other — wrap as method
                members.push(ClassMember::Method(Box::new(stmt)));
            }
        }
    }
    members
}

// ── Decorated ───────────────────────────────────────────────────────────────

fn walk_decorated(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut decorators = Vec::new();
    let mut inner_pairs: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Collect decorators
    let i = 0;
    while i < inner_pairs.len() {
        if inner_pairs[i].as_rule() == Rule::decorator {
            let dec_pair = inner_pairs.remove(i);
            for dp in dec_pair.into_inner() {
                if dp.as_rule() != Rule::NEWLINE {
                    decorators.push(walk_expression(dp)?);
                }
            }
        } else {
            break;
        }
    }

    // Remaining should be the def/class
    if let Some(item) = inner_pairs.into_iter().next() {
        match item.as_rule() {
            Rule::function_def => walk_func_def(item, false, decorators),
            Rule::class_def => walk_class_def(item, decorators),
            Rule::async_stmt => {
                // async def with decorators
                for p in item.into_inner() {
                    match p.as_rule() {
                        Rule::function_def => return walk_func_def(p, true, decorators),
                        Rule::for_stmt => return walk_for(p, true),
                        Rule::with_stmt => return walk_with(p, true),
                        _ => {}
                    }
                }
                Err("Expected def/for/with after async".into())
            }
            other => Err(format!(
                "Expected def/class after decorator, got {:?}",
                other
            )),
        }
    } else {
        Err("Empty decorated statement".into())
    }
}

// ── Async ───────────────────────────────────────────────────────────────────

fn walk_async_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::function_def => return walk_func_def(p, true, Vec::new()),
            Rule::for_stmt => return walk_for(p, true),
            Rule::with_stmt => return walk_with(p, true),
            Rule::async_kw => {}
            _ => {}
        }
    }
    Err("Expected def/for/with after async".into())
}

// ── If ──────────────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let then_body = walk_block(next_rule_any(&mut inner, &[Rule::block])?)?;

    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::elif_clause => {
                let mut ei = p.into_inner();
                let econd = walk_expression(next_meaningful(&mut ei)?)?;
                let ebody = walk_block(next_rule_any(&mut ei, &[Rule::block])?)?;
                elifs.push((econd, ebody));
            }
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

// ── While ───────────────────────────────────────────────────────────────────

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let body = walk_block(next_rule_any(&mut inner, &[Rule::block])?)?;
    let mut else_body = None;
    for p in inner {
        if p.as_rule() == Rule::else_clause {
            let mut ei = p.into_inner();
            else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
        }
    }
    Ok(StmtKind::While {
        cond,
        body,
        else_body,
    })
}

// ── For ─────────────────────────────────────────────────────────────────────

fn walk_for(pair: Pair<Rule>, is_async: bool) -> Result<StmtKind, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Find target_list, expression_list, block, else_clause
    let mut var_names: Vec<String> = Vec::new();
    let mut iter_expr = None;
    let mut body = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::target_list => {
                let text = p.as_str().trim().to_string();
                var_names = text
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();
            }
            Rule::expression_list => {
                if iter_expr.is_none() {
                    iter_expr = Some(walk_expr_list(p)?);
                }
            }
            Rule::block => body = walk_block(p)?,
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            Rule::in_kw | Rule::async_kw => {}
            _ => {}
        }
    }

    // If multiple targets (tuple unpacking: `for i, v in enumerate(...)`),
    // use a temp var and prepend destructuring assignments to the body.
    let var = if var_names.len() > 1 {
        let tmp = "__forin_element".to_string();
        let mut destructure_stmts: Vec<Statement> = Vec::new();
        for (i, name) in var_names.iter().enumerate() {
            // name = __forin_element[i]
            destructure_stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(i as i64)))),
                    null_safe: false,
                }),
            }));
        }
        destructure_stmts.extend(body);
        body = destructure_stmts;
        tmp
    } else {
        var_names.into_iter().next().unwrap_or_default()
    };

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter: iter_expr.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
        body,
        of: true,
        else_body,
        is_async,
    })
}

// ── Try ─────────────────────────────────────────────────────────────────────

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut else_body = None;
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block => {
                if body.is_empty() {
                    body = walk_block(p)?;
                }
            }
            Rule::except_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::block => catch_body = walk_block(cp)?,
                        Rule::identifier => var_name = Some(cp.as_str().to_string()),
                        Rule::as_kw => {}
                        _ => {
                            // Exception type expression
                            types.push(cp.as_str().trim().to_string());
                        }
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
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block {
                        finally = Some(walk_block(fp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body,
        finally,
    })
}

// ── With ────────────────────────────────────────────────────────────────────

fn walk_with(pair: Pair<Rule>, is_async: bool) -> Result<StmtKind, String> {
    let mut items = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::with_item => {
                let mut expr = None;
                let mut var = None;
                for wp in p.into_inner() {
                    match wp.as_rule() {
                        Rule::as_kw => {}
                        Rule::target | Rule::target_list => {
                            var = Some(wp.as_str().trim().to_string())
                        }
                        _ => {
                            if expr.is_none() {
                                expr = Some(walk_expression(wp)?);
                            }
                        }
                    }
                }
                items.push(WithItem {
                    expr: expr.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
                    var,
                });
            }
            Rule::block => body = walk_block(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::With {
        items,
        body,
        is_async,
    })
}

// ── Match ───────────────────────────────────────────────────────────────────

fn walk_match(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut subject = None;
    let mut cases = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression_list => {
                if subject.is_none() {
                    subject = Some(walk_expr_list(p)?);
                }
            }
            Rule::case_clause => {
                let mut pattern = Pattern::Wildcard;
                let mut guard = None;
                let mut body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::pattern | Rule::or_pattern => pattern = walk_pattern(cp)?,
                        Rule::block => body = walk_block(cp)?,
                        _ => {
                            // Guard expression (after "if")
                            if cp.as_rule() != Rule::if_kw {
                                guard = Some(walk_expression(cp)?);
                            }
                        }
                    }
                }
                cases.push(MatchCase {
                    pattern,
                    guard,
                    body,
                });
            }
            Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => {}
            _ => {}
        }
    }

    Ok(StmtKind::MatchStatement {
        subject: subject.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
        cases,
    })
}

fn walk_pattern(pair: Pair<Rule>) -> Result<Pattern, String> {
    match pair.as_rule() {
        Rule::or_pattern => {
            let pats: Vec<Pair<Rule>> = pair.into_inner().collect();
            if pats.len() == 1 {
                walk_pattern(pats.into_iter().next().unwrap())
            } else {
                let patterns = pats
                    .into_iter()
                    .map(walk_pattern)
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(Pattern::Or(patterns))
            }
        }
        Rule::pattern => {
            let inner = pair.into_inner().next();
            match inner {
                Some(p) => walk_pattern(p),
                None => Ok(Pattern::Wildcard),
            }
        }
        Rule::single_pattern => {
            let inner = pair.into_inner().next().ok_or("Empty single_pattern")?;
            walk_pattern(inner)
        }
        Rule::group_pattern => {
            let inner = pair.into_inner().next().ok_or("Empty group_pattern")?;
            walk_pattern(inner)
        }
        Rule::as_pattern => {
            // pattern as name
            let mut inner = pair.into_inner();
            let sub_pattern = walk_pattern(inner.next().ok_or("Missing as_pattern sub-pattern")?)?;
            // skip as_kw
            let name = inner
                .filter(|p| p.as_rule() == Rule::identifier)
                .next()
                .map(|p| p.as_str().to_string());
            Ok(Pattern::As {
                pattern: Some(Box::new(sub_pattern)),
                name,
            })
        }
        Rule::wildcard_pattern => Ok(Pattern::Wildcard),
        Rule::capture_pattern => {
            let name = pair.as_str().to_string();
            Ok(Pattern::As {
                pattern: None,
                name: Some(name),
            })
        }
        Rule::singleton_pattern => {
            let text = pair.as_str().trim();
            let expr = match text {
                "None" => Expression::null(),
                "True" => Expression::bool(true),
                "False" => Expression::bool(false),
                _ => Expression::null(),
            };
            Ok(Pattern::Singleton(expr))
        }
        Rule::literal_pattern => {
            let text = pair.as_str().trim();
            let expr = parse_literal_to_expr(text);
            Ok(Pattern::Value(expr))
        }
        Rule::value_pattern => {
            // Dotted name like module.CONST
            let expr = Expression::new(ExprKind::Ident(pair.as_str().trim().to_string()));
            Ok(Pattern::Value(expr))
        }
        Rule::star_pattern => {
            let name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string());
            Ok(Pattern::Star(name))
        }
        Rule::sequence_pattern => {
            let pats = pair
                .into_inner()
                .map(walk_pattern)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Pattern::Sequence(pats))
        }
        Rule::tuple_pattern => {
            let pats = pair
                .into_inner()
                .map(walk_pattern)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Pattern::Sequence(pats))
        }
        Rule::mapping_pattern => {
            let mut pairs_vec = Vec::new();
            for mp in pair.into_inner() {
                if mp.as_rule() == Rule::mapping_pair {
                    let mut mi = mp.into_inner();
                    let key = walk_expression(mi.next().ok_or("Missing mapping key")?)?;
                    let val = walk_pattern(mi.next().ok_or("Missing mapping pattern")?)?;
                    pairs_vec.push((key, val));
                }
            }
            Ok(Pattern::Mapping(pairs_vec))
        }
        Rule::class_pattern => {
            let mut cls_name = String::new();
            let mut patterns = Vec::new();
            let mut kw_patterns = Vec::new();
            for cp in pair.into_inner() {
                match cp.as_rule() {
                    Rule::identifier => cls_name = cp.as_str().to_string(),
                    Rule::class_pattern_arg => {
                        let mut ai = cp.into_inner();
                        let first = ai.next().ok_or("Empty class_pattern_arg")?;
                        if first.as_rule() == Rule::identifier {
                            // Could be keyword=pattern or just a capture pattern
                            if let Some(second) = ai.next() {
                                // keyword = pattern
                                let name = first.as_str().to_string();
                                let pat = walk_pattern(second)?;
                                kw_patterns.push((name, pat));
                            } else {
                                // Just a pattern (identifier is capture or wildcard)
                                patterns.push(walk_pattern(first)?);
                            }
                        } else {
                            patterns.push(walk_pattern(first)?);
                        }
                    }
                    _ => patterns.push(walk_pattern(cp)?),
                }
            }
            Ok(Pattern::Class {
                cls: Expression::new(ExprKind::Ident(cls_name)),
                patterns,
                kw_patterns,
            })
        }
        Rule::true_kw => Ok(Pattern::Singleton(Expression::bool(true))),
        Rule::false_kw => Ok(Pattern::Singleton(Expression::bool(false))),
        Rule::none_kw => Ok(Pattern::Singleton(Expression::null())),
        other => Err(format!("Unexpected pattern rule: {:?}", other)),
    }
}

// ── Return / Raise ──────────────────────────────────────────────────────────

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| is_expression_rule(p.as_rule()))
        .map(walk_expr_list_or_single)
        .transpose()?;
    Ok(StmtKind::Return(expr))
}

fn walk_raise(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exc = None;
    let mut cause = None;
    let mut saw_from = false;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::from_kw => saw_from = true,
            _ if is_expression_rule(p.as_rule()) => {
                if saw_from {
                    cause = Some(walk_expression(p)?);
                } else {
                    exc = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::Throw { expr: exc, cause })
}

// ── Del / Assert / Global / Nonlocal ────────────────────────────────────────

fn walk_del(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let exprs = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    Ok(StmtKind::Delete(exprs))
}

fn walk_assert(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs: Vec<Expression> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    let msg = if exprs.len() > 1 { exprs.pop() } else { None };
    let test = exprs.into_iter().next().unwrap_or(Expression::bool(false));
    Ok(StmtKind::Assert { test, msg })
}

fn walk_scope_decl(pair: Pair<Rule>, kind: ScopeDeclKind) -> Result<StmtKind, String> {
    let names = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::identifier)
        .map(|p| p.as_str().to_string())
        .collect();
    Ok(StmtKind::ScopeDecl { kind, names })
}

// ── Expression or assignment ────────────────────────────────────────────────

fn walk_expr_or_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .collect();

    if inner.is_empty() {
        return Ok(StmtKind::Empty);
    }

    // Check for augmented assignment op
    let aug_pos = inner
        .iter()
        .position(|p| p.as_rule() == Rule::aug_assign_op);
    if let Some(_pos) = aug_pos {
        let target = walk_expr_list_or_single(inner.remove(0))?;
        let op_str = inner.remove(0).as_str(); // aug_assign_op
        let value = if inner.len() == 1 {
            walk_expr_list_or_single(inner.remove(0))?
        } else {
            walk_remaining_as_expr(&mut inner)?
        };
        let op = match op_str {
            "+=" => CompoundOp::Add,
            "-=" => CompoundOp::Sub,
            "*=" => CompoundOp::Mul,
            "/=" => CompoundOp::Div,
            "//=" => CompoundOp::IDiv,
            "%=" => CompoundOp::Mod,
            "**=" => CompoundOp::Pow,
            "<<=" => CompoundOp::Shl,
            ">>=" => CompoundOp::Shr,
            "|=" => CompoundOp::BitOr,
            "&=" => CompoundOp::BitAnd,
            "^=" => CompoundOp::BitXor,
            "@=" => CompoundOp::Mul, // matmul
            _ => CompoundOp::Add,
        };
        return Ok(StmtKind::CompoundAssign { target, op, value });
    }

    // Check if this has "=" tokens — simple assignment
    // The grammar captures: expression_list ~ ("=" ~ expression_list)+
    // So we may have multiple expression_list separated by = signs
    if inner.len() == 1 {
        let expr = walk_expr_list_or_single(inner.remove(0))?;
        return Ok(StmtKind::Expr(expr));
    }

    // Multiple items => assignment (a = b = c) or annotation (x: int = val)
    // For now, collect all expression_lists and treat last as value
    let mut all_exprs = Vec::new();
    for p in inner {
        if is_expression_rule(p.as_rule()) || p.as_rule() == Rule::expression_list {
            all_exprs.push(walk_expr_list_or_single(p)?);
        }
    }

    if all_exprs.len() >= 2 {
        let value = all_exprs.pop().unwrap();
        // Convert Tuple targets to Destructure for tuple unpacking (x, y = ...)
        let targets = all_exprs
            .into_iter()
            .map(|t| {
                if let ExprKind::Tuple(elems) = &t.kind {
                    let patterns = elems
                        .iter()
                        .map(|e| {
                            if let ExprKind::Ident(name) = &e.kind {
                                ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                            } else {
                                ArrayPatternElem::Hole
                            }
                        })
                        .collect();
                    Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)))
                } else {
                    t
                }
            })
            .collect();
        Ok(StmtKind::Assign { targets, value })
    } else if all_exprs.len() == 1 {
        Ok(StmtKind::Expr(all_exprs.remove(0)))
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── Import ──────────────────────────────────────────────────────────────────

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut imports = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::dotted_as_name {
            let mut name = String::new();
            let mut alias = None;
            for dp in p.into_inner() {
                match dp.as_rule() {
                    Rule::dotted_name => name = dp.as_str().to_string(),
                    Rule::identifier => alias = Some(dp.as_str().to_string()),
                    Rule::as_kw => {}
                    _ => {}
                }
            }
            imports.push((name, alias));
        }
    }

    // For simple `import os`, `import os as operating_system`
    if imports.len() == 1 {
        let (path, alias) = imports.remove(0);
        Ok(Import {
            kind: ImportKind::Simple { path, alias },
            span,
        })
    } else {
        // Multiple: import os, sys — emit first, rest are separate
        let (path, alias) = imports.remove(0);
        Ok(Import {
            kind: ImportKind::Simple { path, alias },
            span,
        })
    }
}

fn walk_import_from(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut level = 0usize;
    let mut module = String::new();
    let mut names = Vec::new();
    let mut is_wildcard = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::import_dots => {
                level = p.as_str().chars().filter(|c| *c == '.').count();
            }
            Rule::dotted_name => module = p.as_str().to_string(),
            Rule::import_names => {
                let text = p.as_str().trim();
                if text == "*" {
                    is_wildcard = true;
                } else {
                    for np in p.into_inner() {
                        if np.as_rule() == Rule::import_as_name {
                            let mut name = String::new();
                            let mut alias = None;
                            for ip in np.into_inner() {
                                match ip.as_rule() {
                                    Rule::identifier => {
                                        if name.is_empty() {
                                            name = ip.as_str().to_string();
                                        } else {
                                            alias = Some(ip.as_str().to_string());
                                        }
                                    }
                                    Rule::as_kw => {}
                                    _ => {}
                                }
                            }
                            names.push(ImportName { name, alias });
                        }
                    }
                }
            }
            _ => {}
        }
    }

    if is_wildcard {
        Ok(Import {
            kind: ImportKind::Wildcard {
                path: module,
                alias: None,
            },
            span,
        })
    } else {
        Ok(Import {
            kind: ImportKind::Named {
                path: module,
                names,
                level,
            },
            span,
        })
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Block parsing
// ════════════════════════════════════════════════════════════════════════════

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => {}
            Rule::simple_stmt_list => {
                for sp in p.into_inner() {
                    let stmt = walk_statement(sp)?;
                    if !matches!(stmt.kind, StmtKind::Empty) {
                        stmts.push(stmt);
                    }
                }
            }
            _ => {
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    stmts.push(stmt);
                }
            }
        }
    }
    Ok(stmts)
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
        // ── Literals ────────────────────────────────────────────────────
        Rule::numeric_literal => parse_number(pair.as_str()),
        Rule::string_literal => {
            let raw = pair.as_str();
            if is_bytes_prefix(raw) {
                Ok(parse_bytes_literal(raw))
            } else {
                Ok(ExprKind::Lit(Literal::Str(parse_python_string(raw))))
            }
        }
        Rule::string_concat => {
            // Implicit string concatenation: "a" "b" → "ab"
            let mut result = String::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::string_literal => result.push_str(&parse_python_string(p.as_str())),
                    Rule::fstring => {
                        // Can't statically concat f-strings; return as interpolation
                        // For now just treat the whole concat as the first piece that's non-trivial
                        return walk_fstring(p);
                    }
                    _ => {}
                }
            }
            Ok(ExprKind::Lit(Literal::Str(result)))
        }
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::none_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::ellipsis_lit => Ok(ExprKind::Lit(Literal::Ellipsis)),
        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),

        // ── Expression wrappers (unwrap single child) ───────────────────
        Rule::expression
        | Rule::named_expr
        | Rule::ternary_expr
        | Rule::or_expr
        | Rule::and_expr
        | Rule::not_expr
        | Rule::comparison
        | Rule::bitor_expr
        | Rule::bitxor_expr
        | Rule::bitand_expr
        | Rule::shift_expr
        | Rule::additive
        | Rule::multiplicative
        | Rule::power
        | Rule::await_expr
        | Rule::unary => walk_infix_or_unwrap(pair),

        Rule::postfix => walk_postfix(pair),
        Rule::primary => walk_primary(pair),
        Rule::expression_list => walk_expr_list_kind(pair),
        Rule::lambda_expr => walk_lambda(pair),
        Rule::yield_expr => walk_yield(pair),
        Rule::star_expr => {
            let inner = pair.into_inner().next().ok_or("Empty star_expr")?;
            Ok(ExprKind::Spread(Box::new(walk_expression(inner)?)))
        }
        Rule::fstring => walk_fstring(pair),

        // List / dict / set inner (when grammar brackets are stripped)
        Rule::list_inner => walk_list_inner(pair),
        Rule::dict_or_set_inner => walk_dict_or_set(pair),
        Rule::comp_for_arg => {
            // Generator expression: expr comp_clause+
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                return Ok(ExprKind::Lit(Literal::Null));
            }
            let element = walk_expression(inner.remove(0))?;
            let generators = inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Generator,
                element: Box::new(element),
                generators,
            })
        }

        // Subscript items (slice)
        Rule::subscript | Rule::subscript_item => walk_subscript_expr(pair),

        Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => Ok(ExprKind::Lit(Literal::Null)),

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Infix / precedence unwrap ───────────────────────────────────────────────

fn walk_infix_or_unwrap(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Single child — unwrap
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    match rule {
        Rule::expression => {
            // Comma expression → sequence/tuple handled at statement level
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let first = inner.remove(0);
                walk_expr_kind(first)
            }
        }
        Rule::named_expr => {
            // target := value
            if inner.len() == 2 {
                let target = walk_expression(inner.remove(0))?;
                let value = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Walrus {
                    target: Box::new(target),
                    value: Box::new(value),
                })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }
        Rule::ternary_expr => {
            // body if_kw test else_kw orelse
            if inner.len() >= 3 {
                let body = walk_expression(inner.remove(0))?;
                // skip if_kw
                let mut rest = inner
                    .into_iter()
                    .filter(|p| p.as_rule() != Rule::if_kw && p.as_rule() != Rule::else_kw);
                let test = walk_expression(rest.next().ok_or("Missing ternary test")?)?;
                let orelse = walk_expression(rest.next().ok_or("Missing ternary else")?)?;
                Ok(ExprKind::Ternary {
                    cond: Box::new(test),
                    then: Box::new(body),
                    else_: Box::new(orelse),
                })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::not_expr => {
            // not_kw ~ not_expr — unary not
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            Ok(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            })
        }
        Rule::comparison => {
            // left (comp_op right)*
            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i < inner.len() {
                let op_pair = &inner[i];
                let op = if op_pair.as_rule() == Rule::comparison_op {
                    let op = parse_comparison_op(op_pair.as_str().trim());
                    i += 1;
                    op
                } else {
                    // Direct expression — shouldn't happen but handle gracefully
                    break;
                };
                if i < inner.len() {
                    let right = walk_expression(inner[i].clone())?;
                    i += 1;
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
                }
            }
            Ok(left.kind)
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::additive => walk_binary_chain_with_ops(inner),
        Rule::multiplicative => walk_python_multiplicative(inner),
        Rule::unary => {
            // unary_op ~ unary
            let op_str = inner[0].as_str().trim();
            let operand = walk_expression(inner.pop().ok_or("Empty unary")?)?;
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand),
            })
        }
        Rule::power => {
            // base ** exponent
            let base = walk_expression(inner.remove(0))?;
            // skip ** op
            let mut rest = inner
                .into_iter()
                .filter(|p| is_expression_rule(p.as_rule()));
            if let Some(exp_pair) = rest.next() {
                let exp = walk_expression(exp_pair)?;
                Ok(ExprKind::Binary {
                    op: BinOp::Pow,
                    left: Box::new(base),
                    right: Box::new(exp),
                })
            } else {
                Ok(base.kind)
            }
        }
        Rule::await_expr => {
            // await_kw ~ unary
            let expr = walk_expression(inner.pop().ok_or("Empty await")?)?;
            Ok(ExprKind::Await(Box::new(expr)))
        }
        _ => {
            if !inner.is_empty() {
                walk_expr_kind(inner.remove(0))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_expression_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            let op = op_fn("");
            left = Expression::new(ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }
    }
    Ok(left.kind)
}

/// Python-specific: `*` is dynamic (str repeat OR numeric mul).
/// Emits Call(__vybe_dynmul, [a, b]) for `*`, delegates others to normal BinOp.
fn walk_python_multiplicative(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                if op_str == "*" {
                    // Dynamic multiply via stdlib
                    let callee = Expression::new(ExprKind::Ident("__vybe_dynmul".into()));
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(callee),
                        args: vec![Argument::positional(left), Argument::positional(right)],
                        optional: false,
                    });
                } else if op_str == "//" {
                    // Python floor division: floor(a / b)
                    let div = Expression::new(ExprKind::Binary {
                        op: BinOp::Div,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
                    let callee = Expression::new(ExprKind::Ident("floor".into()));
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(callee),
                        args: vec![Argument::positional(div)],
                        optional: false,
                    });
                } else {
                    let op = parse_binop(op_str);
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
                }
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
        if is_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim());
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
        } else if is_expression_rule(p.as_rule()) {
            // Operator was merged into the rule text, parse from context
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

// ── Postfix (call, member, subscript chain) ─────────────────────────────────

fn walk_postfix(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty postfix")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() == Rule::postfix_chain {
            // In pest, string literals ("(", ".", "[", "]", ")") are silently consumed
            // in non-atomic rules. So postfix_chain children are just:
            //   call:      call_args? (may be empty for no-arg calls)
            //   member:    identifier
            //   subscript: subscript
            let children: Vec<Pair<Rule>> = chain.into_inner().collect();
            if children.is_empty() {
                // No-arg call: foo()
                // Python `super()` → ExprKind::Super so the compiler's
                // existing super.method() dispatch takes over.
                if matches!(&expr.kind, ExprKind::Ident(n) if n == "super") {
                    expr = Expression::new(ExprKind::Super);
                } else {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: Vec::new(),
                        optional: false,
                    });
                }
            } else {
                let first_child = &children[0];
                match first_child.as_rule() {
                    Rule::call_args => {
                        let args = walk_call_args(children.into_iter().next().unwrap())?;
                        // Python-specific: `delim.join(array)` → swap receiver/arg
                        // so the common compiler sees `array.join(delim)` convention.
                        if let ExprKind::Member {
                            object,
                            field,
                            null_safe,
                        } = &expr.kind
                        {
                            if field == "join" && args.len() == 1 {
                                let delim = object.clone();
                                let array_arg = args.into_iter().next().unwrap().value;
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: Box::new(array_arg),
                                        field: "join".into(),
                                        null_safe: *null_safe,
                                    })),
                                    args: vec![Argument::positional(*delim)],
                                    optional: false,
                                });
                                continue;
                            }
                        }
                        // Python `super(Type, self)` explicit 2-arg form → ExprKind::Super
                        if matches!(&expr.kind, ExprKind::Ident(n) if n == "super") {
                            if args.len() == 0 || args.len() == 2 {
                                expr = Expression::new(ExprKind::Super);
                                continue;
                            }
                        }

                        // Python-specific: rewrite builtins that differ from JS semantics.
                        if let ExprKind::Ident(name) = &expr.kind {
                            match name.as_str() {
                                "divmod" if args.len() == 2 => {
                                    // divmod(a, b) → [a // b, a % b]
                                    let a = args[0].value.clone();
                                    let b = args[1].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![
                                        ArrayElement {
                                            key: None,
                                            spread: false,
                                            by_ref: false,
                                            value: Expression::new(ExprKind::Binary {
                                                op: BinOp::FloorDiv,
                                                left: Box::new(a.clone()),
                                                right: Box::new(b.clone()),
                                            }),
                                        },
                                        ArrayElement {
                                            key: None,
                                            spread: false,
                                            by_ref: false,
                                            value: Expression::new(ExprKind::Binary {
                                                op: BinOp::Mod,
                                                left: Box::new(a),
                                                right: Box::new(b),
                                            }),
                                        },
                                    ]));
                                    continue;
                                }
                                "int" if args.len() == 2 => {
                                    // int(s, base) → parseInt(s, base)
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "parseInt".into(),
                                        ))),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                "isinstance" if args.len() == 2 => {
                                    if let ExprKind::Ident(type_name) = &args[1].value.kind {
                                        if type_name == "int" {
                                            // isinstance(x, int) → typeof x === "number" || typeof x === "boolean"
                                            // because bool is a subtype of int in Python
                                            let x = args[0].value.clone();
                                            expr = Expression::new(ExprKind::Binary {
                                                op: BinOp::Or,
                                                left: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::StrictEq,
                                                    left: Box::new(Expression::new(
                                                        ExprKind::TypeOf(Box::new(x.clone())),
                                                    )),
                                                    right: Box::new(Expression::string("number")),
                                                })),
                                                right: Box::new(Expression::new(
                                                    ExprKind::Binary {
                                                        op: BinOp::StrictEq,
                                                        left: Box::new(Expression::new(
                                                            ExprKind::TypeOf(Box::new(x)),
                                                        )),
                                                        right: Box::new(Expression::string(
                                                            "boolean",
                                                        )),
                                                    },
                                                )),
                                            });
                                            continue;
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args,
                            optional: false,
                        });
                    }
                    Rule::identifier => {
                        let field = first_child.as_str().to_string();
                        if field == "__dict__" {
                            // Python `obj.__dict__` → the object itself.
                            // Vybe stores instance/class properties in Object.properties,
                            // so ARRAY_GET on the object finds the same keys.
                        } else {
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field,
                                null_safe: false,
                            });
                        }
                    }
                    Rule::subscript => {
                        let index = walk_subscript_expr(children.into_iter().next().unwrap())?;
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(Expression::new(index)),
                            null_safe: false,
                        });
                    }
                    _ => {
                        // Fallback: try to walk as expression
                        let val = walk_expression(children.into_iter().next().unwrap())?;
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(val),
                            null_safe: false,
                        });
                    }
                }
            }
        }
    }
    Ok(expr.kind)
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            let mut ci: Vec<Pair<Rule>> = p.into_inner().collect();
            if ci.is_empty() {
                continue;
            }

            let first_text = ci[0].as_str();

            if first_text == "**" {
                // **kwargs
                if ci.len() > 1 {
                    let val = walk_expression(ci.remove(1))?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if first_text == "*" {
                // *args
                if ci.len() > 1 {
                    let val = walk_expression(ci.remove(1))?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if ci.len() >= 2 && ci[0].as_rule() == Rule::identifier {
                // Check if it's keyword=value: identifier followed by expression
                // If there's an "=" between them
                let name = ci[0].as_str().to_string();
                let val = walk_expression(ci.pop().unwrap())?;
                args.push(Argument {
                    value: val,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                });
            } else if ci[0].as_rule() == Rule::comp_for_arg {
                // Generator expression as argument
                let val = walk_expression(ci.remove(0))?;
                args.push(Argument::positional(val));
            } else {
                let val = walk_expression(ci.remove(0))?;
                args.push(Argument::positional(val));
            }
        }
    }
    Ok(args)
}

// ── Subscript ───────────────────────────────────────────────────────────────

fn walk_subscript_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let text = pair.as_str().trim();
    if pair.as_rule() == Rule::subscript_item && text.contains(':') {
        let mut exprs = pair
            .into_inner()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?
            .into_iter();
        let mut parts = text.split(':').map(str::trim);
        let lower = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice lower bound")?)),
        };
        let upper = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice upper bound")?)),
        };
        let step = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice step")?)),
        };
        return Ok(ExprKind::Slice { lower, upper, step });
    }

    let items: Vec<Pair<Rule>> = pair.into_inner().collect();
    if items.len() == 1 {
        return walk_expr_kind(items.into_iter().next().unwrap());
    }
    // Multiple subscript items → tuple
    let exprs = items
        .into_iter()
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    Ok(ExprKind::Tuple(exprs))
}

// ── Primary ─────────────────────────────────────────────────────────────────

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // Multiple children in primary — could be parenthesized expr, list, dict, etc.
    if inner.is_empty() {
        return Ok(ExprKind::Tuple(Vec::new())); // empty tuple ()
    }
    // Check what we have
    let first = &inner[0];
    match first.as_rule() {
        Rule::expression_list => {
            // Parenthesized expression or tuple
            let expr = walk_expr_list(inner.remove(0))?;
            Ok(expr.kind)
        }
        Rule::list_inner => walk_list_inner(inner.remove(0)),
        Rule::dict_or_set_inner => walk_dict_or_set(inner.remove(0)),
        _ => walk_expr_kind(inner.remove(0)),
    }
}

fn walk_list_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.is_empty() {
        return Ok(ExprKind::Array(Vec::new()));
    }

    // Check for comprehension
    let has_comp = inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
    if has_comp {
        let element = walk_expression(inner.remove(0))?;
        let generators = inner
            .into_iter()
            .filter(|p| p.as_rule() == Rule::comp_clause)
            .map(walk_comp_clause)
            .collect::<Result<Vec<_>, _>>()?;
        return Ok(ExprKind::Comprehension {
            kind: ComprehensionKind::List,
            element: Box::new(element),
            generators,
        });
    }

    // Normal list
    let elements = inner
        .into_iter()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            Ok(ArrayElement {
                key: None,
                value: val,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_dict_or_set(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let text = pair.as_str().trim();
    if text.is_empty() {
        return Ok(ExprKind::Object(Vec::new())); // empty dict {}
    }

    let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.is_empty() {
        return Ok(ExprKind::Object(Vec::new()));
    }

    // ── Set comprehension: expression ~ set_comp_or_rest(comp_clause+) ──
    // e.g. {x % 2 for x in range(6)}
    if inner.len() >= 2 && inner[1].as_rule() == Rule::set_comp_or_rest {
        let set_inner: Vec<Pair<Rule>> = inner[1].clone().into_inner().collect();
        let has_comp = set_inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
        if has_comp {
            let element = walk_expression(inner[0].clone())?;
            let generators = set_inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Set,
                element: Box::new(element),
                generators,
            });
        }
    }

    // ── Dict comprehension: expr ~ expr ~ dict_comp_or_rest(comp_clause+) ──
    // e.g. {x: x * x for x in range(4)}
    if inner.len() >= 3 && inner[2].as_rule() == Rule::dict_comp_or_rest {
        let comp_inner: Vec<Pair<Rule>> = inner[2].clone().into_inner().collect();
        let has_comp = comp_inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
        if has_comp
            && is_expression_rule(inner[0].as_rule())
            && is_expression_rule(inner[1].as_rule())
        {
            let key = walk_expression(inner[0].clone())?;
            let val = walk_expression(inner[1].clone())?;
            // Encode key-value as a 2-element array so the compiler can unpack it.
            let element = Expression::new(ExprKind::Array(vec![
                ArrayElement {
                    key: None,
                    spread: false,
                    by_ref: false,
                    value: key,
                },
                ArrayElement {
                    key: None,
                    spread: false,
                    by_ref: false,
                    value: val,
                },
            ]));
            let generators = comp_inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Dict,
                element: Box::new(element),
                generators,
            });
        }
    }

    // ── Dict literal or set literal ──────────────────────────────────────
    let mut is_dict = false;
    for p in &inner {
        match p.as_rule() {
            Rule::dict_comp_or_rest | Rule::dict_rest | Rule::dict_entry => is_dict = true,
            _ => {}
        }
    }

    if is_dict {
        let mut props = Vec::new();
        let mut i = 0;
        while i < inner.len() {
            match inner[i].as_rule() {
                Rule::dict_comp_or_rest | Rule::dict_rest => {
                    for de in inner[i].clone().into_inner() {
                        if de.as_rule() == Rule::dict_entry {
                            let is_spread = de.as_str().trim_start().starts_with("**");
                            let entry_inner: Vec<Pair<Rule>> = de.into_inner().collect();
                            if is_spread {
                                if let Some(expr) = entry_inner.first() {
                                    props.push(ObjectProperty::Spread(walk_expression(
                                        expr.clone(),
                                    )?));
                                }
                            } else if entry_inner.len() >= 2 {
                                let key = walk_expression(entry_inner[0].clone())?;
                                let val = walk_expression(entry_inner[1].clone())?;
                                props.push(ObjectProperty::KeyValue { key, value: val });
                            }
                        }
                    }
                }
                _ if is_expression_rule(inner[i].as_rule()) => {
                    let key = walk_expression(inner[i].clone())?;
                    if i == 0 && text.starts_with("**") {
                        props.push(ObjectProperty::Spread(key));
                        i += 1;
                        continue;
                    }
                    i += 1;
                    if i < inner.len() && is_expression_rule(inner[i].as_rule()) {
                        let val = walk_expression(inner[i].clone())?;
                        props.push(ObjectProperty::KeyValue { key, value: val });
                    }
                }
                _ => {}
            }
            i += 1;
        }
        return Ok(ExprKind::Object(props));
    }

    // Set literal: {1, 2, 3}
    let mut elements = Vec::new();
    for item in inner {
        match item.as_rule() {
            rule if is_expression_rule(rule) => elements.push(walk_expression(item)?),
            Rule::set_comp_or_rest => {
                for part in item.into_inner() {
                    if is_expression_rule(part.as_rule()) {
                        elements.push(walk_expression(part)?);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(ExprKind::Set(elements))
}

fn walk_comp_clause(pair: Pair<Rule>) -> Result<ComprehensionGen, String> {
    let mut target = Expression::new(ExprKind::Ident("_".into()));
    let mut iter = Expression::new(ExprKind::Lit(Literal::Null));
    let mut conditions = Vec::new();
    let mut is_async = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::async_kw => is_async = true,
            Rule::target_list => {
                target = Expression::new(ExprKind::Ident(p.as_str().trim().to_string()));
            }
            Rule::in_kw => {}
            Rule::comp_if => {
                for ci in p.into_inner() {
                    if is_expression_rule(ci.as_rule()) {
                        conditions.push(walk_expression(ci)?);
                    }
                }
            }
            _ if is_expression_rule(p.as_rule()) => {
                iter = walk_expression(p)?;
            }
            _ => {}
        }
    }

    Ok(ComprehensionGen {
        target,
        iter,
        conditions,
        is_async,
    })
}

// ── Lambda ──────────────────────────────────────────────────────────────────

fn walk_lambda(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body_expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::lambda_params => {
                for lp in p.into_inner() {
                    if lp.as_rule() == Rule::lambda_param {
                        let mut name = String::new();
                        let mut default = None;
                        let mut is_rest = false;
                        let mut is_kwargs = false;
                        for c in lp.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ if c.as_str() == "**" => is_kwargs = true,
                                _ if c.as_str() == "*" => is_rest = true,
                                _ => default = Some(walk_expression(c)?),
                            }
                        }
                        if !name.is_empty() {
                            params.push(Param {
                                name,
                                type_hint: None,
                                is_optional: default.is_some(),
                                default,
                                pass_by: PassBy::Value,
                                is_rest,
                                is_kwargs,
                                is_nullable: false,
                            });
                        }
                    }
                }
            }
            _ if is_expression_rule(p.as_rule()) => {
                body_expr = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(Box::new(body_expr.unwrap_or(Expression::null()))),
        is_async: false,
        captures: Vec::new(),
    })
}

// ── Yield ───────────────────────────────────────────────────────────────────

fn walk_yield(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut is_from = false;
    let mut expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::yield_kw => {}
            Rule::yield_from_kw => is_from = true,
            _ if is_expression_rule(p.as_rule()) => expr = Some(walk_expression(p)?),
            Rule::expression_list => expr = Some(walk_expr_list(p)?),
            _ => {}
        }
    }

    if is_from {
        Ok(ExprKind::YieldFrom(Box::new(
            expr.unwrap_or(Expression::null()),
        )))
    } else {
        Ok(ExprKind::Yield(expr.map(Box::new)))
    }
}

// ── F-string ────────────────────────────────────────────────────────────────

fn walk_fstring(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::fstring_start | Rule::fstring_end => {}
            Rule::fstring_text => {
                parts.push(InterpolPart::Text(p.as_str().to_string()));
            }
            Rule::fstring_escaped_brace => {
                let text = if p.as_str().starts_with('{') {
                    "{"
                } else {
                    "}"
                };
                parts.push(InterpolPart::Text(text.into()));
            }
            Rule::fstring_expr => {
                for fp in p.into_inner() {
                    if is_expression_rule(fp.as_rule()) {
                        parts.push(InterpolPart::Expr(walk_expression(fp)?));
                    }
                }
            }
            _ => {}
        }
    }
    Ok(ExprKind::Interpolation(parts))
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn walk_expr_list(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expression(inner.remove(0))
    } else if inner.is_empty() {
        Ok(Expression::new(ExprKind::Tuple(Vec::new())))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(Expression::with_span(ExprKind::Tuple(exprs), span))
    }
}

fn walk_expr_list_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expr_kind(inner.into_iter().next().unwrap())
    } else if inner.is_empty() {
        Ok(ExprKind::Tuple(Vec::new()))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(ExprKind::Tuple(exprs))
    }
}

fn walk_expr_list_or_single(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::expression_list {
        walk_expr_list(pair)
    } else {
        walk_expression(pair)
    }
}

fn walk_remaining_as_expr(items: &mut Vec<Pair<Rule>>) -> Result<Expression, String> {
    if items.len() == 1 {
        walk_expression(items.remove(0))
    } else {
        walk_expression(items.remove(0))
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32,
    }
}

fn next_meaningful<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        match p.as_rule() {
            Rule::NEWLINE
            | Rule::INDENT
            | Rule::DEDENT
            | Rule::in_kw
            | Rule::as_kw
            | Rule::async_kw => continue,
            _ => return Ok(p),
        }
    }
    Err("No more meaningful pairs".into())
}

fn next_rule_any<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rules: &[Rule],
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if rules.contains(&p.as_rule()) {
            return Ok(p);
        }
    }
    Err(format!("Expected one of {:?}", rules))
}

fn is_expression_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::expression
            | Rule::expression_list
            | Rule::named_expr
            | Rule::ternary_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::power
            | Rule::await_expr
            | Rule::postfix
            | Rule::primary
            | Rule::lambda_expr
            | Rule::yield_expr
            | Rule::star_expr
            | Rule::fstring
            | Rule::numeric_literal
            | Rule::string_literal
            | Rule::string_concat
            | Rule::identifier
            | Rule::true_kw
            | Rule::false_kw
            | Rule::none_kw
            | Rule::ellipsis_lit
            | Rule::subscript
            | Rule::subscript_item
            | Rule::comp_for_arg
    )
}

fn is_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::multiplicative_op
            | Rule::shift_op
            | Rule::comparison_op
            | Rule::aug_assign_op
    )
}

fn parse_number(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    // Complex numbers (j suffix)
    if s.ends_with('j') || s.ends_with('J') {
        let num_str = &s[..s.len() - 1];
        let val: f64 = num_str.parse().unwrap_or(0.0);
        return Ok(ExprKind::Lit(Literal::Float(val)));
    }
    if s.contains('.')
        || (s.contains('e') || s.contains('E')) && !s.starts_with("0x") && !s.starts_with("0X")
    {
        Ok(ExprKind::Lit(Literal::Float(
            s.parse().map_err(|e| format!("{}", e))?,
        )))
    } else if s.starts_with("0x") || s.starts_with("0X") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 16).unwrap_or(0),
        )))
    } else if s.starts_with("0o") || s.starts_with("0O") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 8).unwrap_or(0),
        )))
    } else if s.starts_with("0b") || s.starts_with("0B") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 2).unwrap_or(0),
        )))
    } else {
        Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
    }
}

fn is_bytes_prefix(s: &str) -> bool {
    let lc = s.to_ascii_lowercase();
    lc.starts_with("b'")
        || lc.starts_with("b\"")
        || lc.starts_with("rb'")
        || lc.starts_with("rb\"")
        || lc.starts_with("br'")
        || lc.starts_with("br\"")
}

fn parse_bytes_literal(s: &str) -> ExprKind {
    let content = parse_python_string(s);
    let elements = content
        .bytes()
        .map(|b| ArrayElement {
            key: None,
            spread: false,
            by_ref: false,
            value: Expression::new(ExprKind::Lit(Literal::Int(b as i64))),
        })
        .collect();
    ExprKind::Array(elements)
}

fn parse_python_string(s: &str) -> String {
    let mut s = s;
    // Strip prefix (r, b, rb, u, etc.)
    let prefixes = [
        "rb", "Rb", "rB", "RB", "br", "bR", "Br", "BR", "r", "R", "b", "B", "u", "U",
    ];
    for prefix in &prefixes {
        if s.starts_with(prefix) {
            s = &s[prefix.len()..];
            break;
        }
    }
    // Strip quotes
    if s.starts_with("\"\"\"") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with("'''") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with('"') {
        s = &s[1..s.len() - 1];
    } else if s.starts_with('\'') {
        s = &s[1..s.len() - 1];
    }
    // Basic escape processing
    s.replace("\\n", "\n")
        .replace("\\t", "\t")
        .replace("\\r", "\r")
        .replace("\\\\", "\\")
        .replace("\\'", "'")
        .replace("\\\"", "\"")
}

fn parse_comparison_op(s: &str) -> BinOp {
    match s {
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "in" => BinOp::In,
        "is" => BinOp::Is,
        _ if s.contains("not") && s.contains("in") => BinOp::NotIn,
        _ if s.contains("is") && s.contains("not") => BinOp::IsNot,
        _ => BinOp::Eq,
    }
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "//" => BinOp::FloorDiv,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "@" => BinOp::MatMul,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn parse_literal_to_expr(text: &str) -> Expression {
    if let Ok(n) = text.parse::<i64>() {
        Expression::int(n)
    } else if let Ok(f) = text.parse::<f64>() {
        Expression::float(f)
    } else {
        Expression::string(text)
    }
}
