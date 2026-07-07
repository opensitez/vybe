use super::{PythonParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::collections::HashMap;

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

    apply_float_var_repr(&mut body, &mut HashMap::new());

    // Prepend the bytes-repr source helper when the program uses bytes, so
    // `b'…'` display resolves to a real `__vybe_bytes_repr` function.
    if source_uses_bytes(source) {
        let mut prelude = parse_python_prelude(BYTES_REPR_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Python,
        body,
        imports,
    })
}

/// Heuristic: does the source reference bytes at all? Only gates whether the
/// repr helper is injected — a false positive just adds an unused function.
fn source_uses_bytes(source: &str) -> bool {
    source.contains("b'")
        || source.contains("b\"")
        || source.contains("B'")
        || source.contains("B\"")
        || source.contains("bytes(")
        || source.contains(".encode(")
        || source.contains(".decode(")
}

/// Parse a Python source prelude into top-level statements. Errors yield `[]`
/// so a prelude problem can never break user compilation.
fn parse_python_prelude(src: &str) -> Vec<Statement> {
    let preprocessed = preprocess_indentation(src);
    let Ok(pairs) = PythonParser::parse(Rule::program, &preprocessed) else {
        return Vec::new();
    };
    let mut body = Vec::new();
    let mut imports = Vec::new();
    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for pair in top.into_inner() {
                    match pair.as_rule() {
                        Rule::EOI | Rule::NEWLINE => continue,
                        _ => {
                            let _ = walk_stmt_into(pair, &mut body, &mut imports);
                        }
                    }
                }
            }
            Rule::EOI => continue,
            _ => {
                let _ = walk_stmt_into(top, &mut body, &mut imports);
            }
        }
    }
    body
}

/// Python source for `__vybe_bytes_repr(int_array) -> "b'…'"`. Escape fragments
/// are built from `chr(92)` (backslash) rather than backslash string literals,
/// which the Python string-escape lowering mishandles.
const BYTES_REPR_PRELUDE: &str = r#"
def __vybe_bytes_repr(a):
    bs = chr(92)
    hexd = "0123456789abcdef"
    r = "b'"
    for b in a:
        if b == 9:
            r += bs + "t"
        elif b == 10:
            r += bs + "n"
        elif b == 13:
            r += bs + "r"
        elif b == 92:
            r += bs + bs
        elif b == 39:
            r += bs + "'"
        elif 32 <= b <= 126:
            r += chr(b)
        else:
            r += bs + "x" + hexd[b >> 4] + hexd[b & 15]
    return r + "'"
"#;

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
    // Any function containing `yield` is a true lazy generator — compiled
    // through the shared stack-switching machinery (`generators.rs`), exactly
    // like JavaScript. No eager list materialization (that hung on `while True`
    // generators and was semantically eager).
    let has_yield = body_has_yield(&body);

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
        is_generator: has_yield,
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
                        // Grammar: `identifier (":" expr)? ("=" expr)?`. pest
                        // drops the `:`/`=` tokens, so with a single expression
                        // we can't tell a type annotation (`x: int`) from a
                        // default (`x = 2`) by rule alone — disambiguate from
                        // the source text (does the part after the name start
                        // with `=`?). With two expressions it's `type = default`.
                        let param_text = item.as_str().trim_start().to_string();
                        let mut name = String::new();
                        let mut default = None;
                        let mut type_hint = None;
                        let mut exprs = Vec::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier && name.is_empty() {
                                name = c.as_str().to_string();
                            } else {
                                exprs.push(c);
                            }
                        }
                        let after_name = param_text
                            .strip_prefix(name.as_str())
                            .unwrap_or("")
                            .trim_start();
                        match exprs.len() {
                            2 => {
                                type_hint = Some(exprs[0].as_str().to_string());
                                default = Some(walk_expression(exprs.remove(1))?);
                            }
                            1 => {
                                if after_name.starts_with('=') {
                                    default = Some(walk_expression(exprs.remove(0))?);
                                } else {
                                    type_hint = Some(exprs[0].as_str().to_string());
                                }
                            }
                            _ => {}
                        }
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

/// Per-parse counter that keeps desugared `while…else` break-flags unique so
/// nested loops don't share (and clobber) one flag.
static WHILE_ELSE_COUNTER: std::sync::atomic::AtomicUsize =
    std::sync::atomic::AtomicUsize::new(0);

/// Rewrite loop-level `break` statements in `stmts` to set `flag = True` before
/// breaking, so a desugared `while…else` can distinguish a break-exit from a
/// normal exit. Recurses through non-loop containers (if/try/with/block) but NOT
/// into nested loops or function/class bodies — their `break`s target
/// themselves, not this loop.
fn mark_loop_break_sets_flag(stmts: Vec<Statement>, flag: &str) -> Vec<Statement> {
    let mut out = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        match stmt.kind {
            StmtKind::Break(_) => {
                out.push(Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(flag)],
                    value: Expression::bool(true),
                }));
                out.push(stmt);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                out.push(Statement::new(StmtKind::If {
                    cond,
                    then_body: mark_loop_break_sets_flag(then_body, flag),
                    elifs: elifs
                        .into_iter()
                        .map(|(c, b)| (c, mark_loop_break_sets_flag(b, flag)))
                        .collect(),
                    else_body: else_body.map(|b| mark_loop_break_sets_flag(b, flag)),
                }));
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                out.push(Statement::new(StmtKind::Try {
                    body: mark_loop_break_sets_flag(body, flag),
                    catches: catches
                        .into_iter()
                        .map(|mut c| {
                            c.body = mark_loop_break_sets_flag(c.body, flag);
                            c
                        })
                        .collect(),
                    else_body: else_body.map(|b| mark_loop_break_sets_flag(b, flag)),
                    finally: finally.map(|b| mark_loop_break_sets_flag(b, flag)),
                }));
            }
            StmtKind::With {
                items,
                body,
                is_async,
            } => {
                out.push(Statement::new(StmtKind::With {
                    items,
                    body: mark_loop_break_sets_flag(body, flag),
                    is_async,
                }));
            }
            StmtKind::Block(b) => {
                out.push(Statement::new(StmtKind::Block(mark_loop_break_sets_flag(
                    b, flag,
                ))));
            }
            _ => out.push(stmt),
        }
    }
    out
}

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

    let Some(else_stmts) = else_body else {
        return Ok(StmtKind::While {
            cond,
            body,
            else_body: None,
        });
    };

    // Python `while C: BODY else: ELSE` runs ELSE only on a NORMAL exit
    // (condition false), never on `break`. The shared While emitter runs
    // else_body unconditionally, so normalize into common-AST primitives that
    // route through the common loop emitter (loops.rs), the same plain-`while`
    // path every language uses:
    //   __while_else_N = False
    //   while C: BODY'                 (loop-level break → __while_else_N = True; break)
    //   if not __while_else_N: ELSE
    let n = WHILE_ELSE_COUNTER.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let flag = format!("__while_else_{n}");
    let body = mark_loop_break_sets_flag(body, &flag);
    let flag_init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&flag)],
        value: Expression::bool(false),
    });
    let while_stmt = Statement::new(StmtKind::While {
        cond,
        body,
        else_body: None,
    });
    let else_guard = Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::ident(&flag)),
        }),
        then_body: else_stmts,
        elifs: Vec::new(),
        else_body: None,
    });
    Ok(StmtKind::Block(vec![flag_init, while_stmt, else_guard]))
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

    // `del obj[key]` (single, non-slice) → `obj.pop(key)`. Python `del` and
    // `.pop()` remove identically (both raise on a missing key/index), and
    // `.pop()` already works for dicts AND lists — whereas `StmtKind::Delete`'s
    // dict branch (`dict::emit_method_delete`) is broken on dict literals (they
    // don't populate the `__keys` array it relies on). Slices, bare names, and
    // multi-target dels keep the existing Delete path.
    if let [target] = exprs.as_slice() {
        if let ExprKind::Index { object, index, .. } = &target.kind {
            if !matches!(index.kind, ExprKind::Slice { .. } | ExprKind::Range { .. }) {
                let pop = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: "pop".into(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false,
                });
                return Ok(StmtKind::Expr(pop));
            }
        }
    }
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
        // `+=` / `*=` use Python's dynamic add/mul (list concat/repeat, string
        // ops), so lower to `target = __pyadd__(target, value)` — the numeric
        // CompoundAssign path coerces operands to f64 and traps on lists.
        if op_str == "+=" || op_str == "*=" {
            let helper = if op_str == "+=" { "__pyadd__" } else { "__pymul__" };
            let combined = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                args: vec![
                    Argument::positional(target.clone()),
                    Argument::positional(value),
                ],
                optional: false,
            });
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value: combined,
            });
        }
        // `-=`/`|=`/`&=`/`^=` lower to `x = x <binop> v` so the polymorphic
        // binary operator handles sets (difference/union/intersection/symmetric
        // difference) as well as integer bitwise / numeric subtraction — the
        // numeric CompoundAssign path only does the arithmetic case.
        let set_binop = match op_str {
            "-=" => Some(BinOp::Sub),
            "|=" => Some(BinOp::BitOr),
            "&=" => Some(BinOp::BitAnd),
            "^=" => Some(BinOp::BitXor),
            // `//=`/`%=` too: binary `//`/`%` use Python floor/mod semantics
            // (round toward -inf, mod follows divisor sign); the CompoundAssign
            // path truncates toward zero.
            "//=" => Some(BinOp::FloorDiv),
            "%=" => Some(BinOp::Mod),
            _ => None,
        };
        if let Some(op) = set_binop {
            let combined = Expression::new(ExprKind::Binary {
                op,
                left: Box::new(target.clone()),
                right: Box::new(value),
            });
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value: combined,
            });
        }
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
                    let patterns = elems.iter().map(expr_to_array_pattern_elem).collect();
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
            // not_kw ~ not_expr — unary not. Lower to `False if bool(x) else True`
            // so Python truthiness applies (empty list/dict/str are falsy) and we
            // route through the working `bool()` / conditional path rather than
            // `emit_dyn_not`, which uses JS truthiness (arrays are always truthy).
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            // The conditional's own condition already applies Python truthiness
            // (`if []:` is falsy), so use the operand directly as the condition.
            Ok(ExprKind::Ternary {
                cond: Box::new(operand),
                then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                else_: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            })
        }
        Rule::comparison => {
            // left (comp_op right)* — Python chains: a < b < c → a < b and b < c
            let mut operands: Vec<Expression> = vec![walk_expression(inner.remove(0))?];
            let mut comparisons: Vec<(BinOp, Expression)> = Vec::new();
            let mut i = 0;
            while i < inner.len() {
                let op_pair = &inner[i];
                let op = if op_pair.as_rule() == Rule::comparison_op {
                    let op = parse_comparison_op(op_pair.as_str().trim());
                    i += 1;
                    op
                } else {
                    break;
                };
                if i < inner.len() {
                    let right = walk_expression(inner[i].clone())?;
                    i += 1;
                    operands.push(right.clone());
                    comparisons.push((op, right));
                }
            }

            if comparisons.len() <= 1 {
                let mut left = operands.remove(0);
                for (op, right) in comparisons {
                    // Normalize `x in {set}` → `{set}.has(x)`
                    if matches!(op, BinOp::In | BinOp::NotIn)
                        && matches!(right.kind, ExprKind::Set(_))
                    {
                        let has_call = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(right),
                                field: "has".into(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(left.clone())],
                            optional: false,
                        });
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(has_call),
                            })
                        } else {
                            has_call
                        };
                    } else if matches!(op, BinOp::In | BinOp::NotIn) {
                        // `x in y` — polymorphic membership (string substring /
                        // list element / dict key). Route to the Python adapter
                        // `__py_contains__(y, x)` rather than the shared
                        // `BinOp::In`, whose runtime array-classification
                        // mis-sends plain objects to `Array.includes`.
                        let contains = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(
                                "__py_contains__".into(),
                            ))),
                            args: vec![
                                Argument::positional(right),
                                Argument::positional(left.clone()),
                            ],
                            optional: false,
                        });
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(contains),
                            })
                        } else {
                            contains
                        };
                    } else {
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        });
                    }
                }
                Ok(left.kind)
            } else {
                let mut result = Expression::new(ExprKind::Binary {
                    op: comparisons[0].0,
                    left: Box::new(operands[0].clone()),
                    right: Box::new(operands[1].clone()),
                });
                for j in 1..comparisons.len() {
                    let pairwise = Expression::new(ExprKind::Binary {
                        op: comparisons[j].0,
                        left: Box::new(operands[j].clone()),
                        right: Box::new(operands[j + 1].clone()),
                    });
                    result = Expression::new(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(result),
                        right: Box::new(pairwise),
                    });
                }
                Ok(result.kind)
            }
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::additive => walk_python_additive(inner),
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
/// Python `+` routes through `__pyadd__` builtin (emitter adapter handles
/// array concat vs string concat vs numeric add). `-` is always numeric.
fn walk_python_additive(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
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
                if op_str == "+" {
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("__pyadd__".into()))),
                        args: vec![Argument::positional(left), Argument::positional(right)],
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
                    // Route through __pymul__ which handles array repeat + string repeat + numeric
                    let callee = Expression::new(ExprKind::Ident("__pymul__".into()));
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(callee),
                        args: vec![Argument::positional(left), Argument::positional(right)],
                        optional: false,
                    });
                } else if op_str == "//" {
                    left = Expression::new(ExprKind::Binary {
                        op: BinOp::FloorDiv,
                        left: Box::new(left),
                        right: Box::new(right),
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
                } else if matches!(&expr.kind, ExprKind::Ident(n) if n == "print") {
                    // Bare `print()` still needs the [sep, end] convention so
                    // the emitter prints the default line terminator.
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("print".into()))),
                        args: normalize_python_print_args(Vec::new()),
                        optional: false,
                    });
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
                            if field == "sort"
                                && args.iter().any(|a| a.name.as_deref() == Some("reverse"))
                            {
                                // arr.sort(reverse=True) → arr.sort(); arr.reverse()
                                let sort_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "sort".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![],
                                    optional: false,
                                });
                                let reverse_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "reverse".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![],
                                    optional: false,
                                });
                                // Chain: sort then reverse. Use comma expression or sequence.
                                // Emit sort as statement, then reverse
                                expr = Expression::new(ExprKind::Sequence(vec![
                                    sort_call,
                                    reverse_call,
                                ]));
                                continue;
                            }
                            if field == "count" && args.len() == 1 {
                                // arr.count(x) → arr.filter(e => e === x).length
                                let needle = args.into_iter().next().unwrap().value;
                                let param = Param {
                                    name: "__e".into(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                };
                                let filter_fn = Expression::new(ExprKind::Lambda {
                                    params: vec![param],
                                    body: LambdaBody::Expr(Box::new(Expression::new(
                                        ExprKind::Binary {
                                            op: BinOp::StrictEq,
                                            left: Box::new(Expression::new(ExprKind::Ident(
                                                "__e".into(),
                                            ))),
                                            right: Box::new(needle),
                                        },
                                    ))),
                                    is_async: false,
                                    captures: vec![],
                                });
                                let filter_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "filter".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![Argument::positional(filter_fn)],
                                    optional: false,
                                });
                                expr = Expression::new(ExprKind::Member {
                                    object: Box::new(filter_call),
                                    field: "length".into(),
                                    null_safe: false,
                                });
                                continue;
                            }
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
                                "print" => {
                                    // Reshape to the emitter convention
                                    // [sep, end, items…]; sep/end kwargs are
                                    // pulled out of the positional list.
                                    let new_args = normalize_python_print_args(args);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "print".into(),
                                        ))),
                                        args: new_args,
                                        optional: false,
                                    });
                                    continue;
                                }
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
                                "bool" if args.len() == 1 => {
                                    // bool(x) → x ? True : False → ternary
                                    let x = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(x),
                                        then: Box::new(Expression::bool(true)),
                                        else_: Box::new(Expression::bool(false)),
                                    });
                                    continue;
                                }
                                "bool" if args.is_empty() => {
                                    expr = Expression::bool(false);
                                    continue;
                                }
                                "list" if args.len() == 1 => {
                                    // list(iterable) → [...iterable]
                                    let iterable = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: iterable,
                                    }]));
                                    continue;
                                }
                                "list" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Array(vec![]));
                                    continue;
                                }
                                "tuple" if args.len() == 1 => {
                                    // tuple(iterable) → [...iterable]
                                    let iterable = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: iterable,
                                    }]));
                                    continue;
                                }
                                "tuple" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Array(vec![]));
                                    continue;
                                }
                                // str(x) is left as a plain call so it routes
                                // through the profile → `common:python.str`
                                // (emit_py_repr), which applies Python repr
                                // semantics: True/False/None, [.., ..] lists,
                                // {'k': v} dicts, single-quoted nested strings.
                                "dict" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Object(vec![]));
                                    continue;
                                }
                                "sum" if !args.is_empty() && args[0].name.is_none() => {
                                    // sum(iterable[, start]) — drain the iterable
                                    // (generators/ranges) via spread first.
                                    let mut new_args = args;
                                    let it = new_args[0].value.clone();
                                    new_args[0].value = spread_iterable_expr(it);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "sum".into(),
                                        ))),
                                        args: new_args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                "min" | "max" | "any" | "all"
                                    if args.len() == 1 && args[0].name.is_none() =>
                                {
                                    // Single-iterable form: drain via spread so
                                    // the array-based builtin sees a sequence.
                                    let n = name.to_string();
                                    let it = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(n))),
                                        args: vec![Argument::positional(spread_iterable_expr(it))],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "filter"
                                    if args.len() == 2
                                        && matches!(
                                            args[0].value.kind,
                                            ExprKind::Lit(Literal::Null)
                                        ) =>
                                {
                                    // filter(None, iter) → keep truthy elements
                                    // (identity predicate `lambda __e: __e`).
                                    let ident = Expression::new(ExprKind::Lambda {
                                        params: vec![Param {
                                            name: "__e".into(),
                                            type_hint: None,
                                            default: None,
                                            pass_by: PassBy::Value,
                                            is_rest: false,
                                            is_kwargs: false,
                                            is_optional: false,
                                            is_nullable: false,
                                        }],
                                        body: LambdaBody::Expr(Box::new(Expression::new(
                                            ExprKind::Ident("__e".into()),
                                        ))),
                                        is_async: false,
                                        captures: vec![],
                                    });
                                    let mut new_args = args;
                                    new_args[0] = Argument::positional(ident);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "filter".into(),
                                        ))),
                                        args: new_args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sorted" if args.len() >= 1 => {
                                    // sorted(iterable) → [...iterable].sort()
                                    // sorted(iterable, reverse=True) → [...iterable].sort().reverse()
                                    let iterable = args[0].value.clone();
                                    let has_reverse =
                                        args.iter().any(|a| a.name.as_deref() == Some("reverse"));
                                    let spread_array =
                                        Expression::new(ExprKind::Array(vec![ArrayElement {
                                            key: None,
                                            spread: true,
                                            by_ref: false,
                                            value: iterable,
                                        }]));
                                    let sorted = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(spread_array),
                                            field: "sort".into(),
                                            null_safe: false,
                                        })),
                                        args: vec![],
                                        optional: false,
                                    });
                                    expr = if has_reverse {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(sorted),
                                                field: "reverse".into(),
                                                null_safe: false,
                                            })),
                                            args: vec![],
                                            optional: false,
                                        })
                                    } else {
                                        sorted
                                    };
                                    continue;
                                }
                                "set" => {
                                    // set() → new Set(), set(iter) → new Set(iter)
                                    expr = Expression::new(ExprKind::New {
                                        class: Box::new(Expression::new(ExprKind::Ident(
                                            "Set".into(),
                                        ))),
                                        args,
                                    });
                                    continue;
                                }
                                "frozenset" => {
                                    expr = Expression::new(ExprKind::New {
                                        class: Box::new(Expression::new(ExprKind::Ident(
                                            "Set".into(),
                                        ))),
                                        args,
                                    });
                                    continue;
                                }
                                "round" if args.len() == 2 => {
                                    // round(x, n) → Math.round(x * 10**n) / 10**n
                                    let x = args[0].value.clone();
                                    let n = args[1].value.clone();
                                    let factor = Expression::new(ExprKind::Binary {
                                        op: BinOp::Pow,
                                        left: Box::new(Expression::int(10)),
                                        right: Box::new(n),
                                    });
                                    let scaled = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mul,
                                        left: Box::new(x),
                                        right: Box::new(factor.clone()),
                                    });
                                    let rounded = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "round".into(),
                                        ))),
                                        args: vec![Argument::positional(scaled)],
                                        optional: false,
                                    });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Div,
                                        left: Box::new(rounded),
                                        right: Box::new(factor),
                                    });
                                    continue;
                                }
                                "pow" if args.len() == 3 => {
                                    // pow(base, exp, mod) → pow(base, exp) % mod
                                    let base = args[0].value.clone();
                                    let exp = args[1].value.clone();
                                    let modulus = args[2].value.clone();
                                    let power = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "pow".into(),
                                        ))),
                                        args: vec![
                                            Argument::positional(base),
                                            Argument::positional(exp),
                                        ],
                                        optional: false,
                                    });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mod,
                                        left: Box::new(power),
                                        right: Box::new(modulus),
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }
                        // Python `gen.throw(ExcClass)` instantiates the class so
                        // the generator's `except` matches an instance (like
                        // `raise ExcClass`). Wrap a bare uppercase-Ident arg.
                        let args = if matches!(&expr.kind, ExprKind::Member { field, .. } if field == "throw")
                            && args.first().is_some_and(|a| {
                                a.name.is_none()
                                    && matches!(&a.value.kind, ExprKind::Ident(n)
                                        if n.chars().next().is_some_and(|c| c.is_ascii_uppercase()))
                            }) {
                            let mut new_args = args;
                            let cls = new_args[0].value.clone();
                            new_args[0].value = Expression::new(ExprKind::Call {
                                callee: Box::new(cls),
                                args: vec![],
                                optional: false,
                            });
                            new_args
                        } else {
                            args
                        };
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
                        let index = Expression::new(walk_subscript_expr(
                            children.into_iter().next().unwrap(),
                        )?);
                        let index = python_index_operand(&expr, index);
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(index),
                            null_safe: false,
                        });
                    }
                    _ => {
                        // Fallback: try to walk as expression
                        let val = walk_expression(children.into_iter().next().unwrap())?;
                        let val = python_index_operand(&expr, val);
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

/// Wrap `value` in `[*value]` so a lazy iterable (generator, range, map, …) is
/// drained into an array via the shared `generators.rs` spread machinery before
/// an array/set-based builtin consumes it. Keeps generators cross-language
/// compatible — the drain is the same one JS spread uses.
fn spread_iterable_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Array(vec![ArrayElement {
        key: None,
        spread: true,
        by_ref: false,
        value,
    }]))
}

/// Normalize Python `print(...)` arguments to the emitter convention
/// `[sep, end, items…]`. The `sep`/`end` keyword args override the defaults
/// (`" "` / `"\n"`); Python's `file`/`flush` keywords are accepted and ignored.
/// Positional items (including `*spread`) keep their original `Argument`.
/// Math functions that return a Python `float` (used by `expr_is_python_float`).
const FLOAT_MATH_FNS: &[&str] = &[
    "sin", "cos", "tan", "asin", "acos", "atan", "atan2", "sinh", "cosh", "tanh", "asinh", "acosh",
    "atanh", "sqrt", "pow", "exp", "log", "log2", "log10", "log1p", "expm1", "cbrt", "degrees",
    "radians", "hypot", "fabs", "fmod", "copysign", "remainder", "dist", "fsum", "gamma", "lgamma",
    "erf", "erfc", "ldexp",
];

/// True when an expression is *statically* a Python `float` — a float literal,
/// true division (`/`), `float()`, a float-returning `math.*` call, unary minus
/// of a float, or arithmetic where an operand is a float. Deliberately
/// conservative: never assumes a bare variable or unknown call is a float
/// (that would be the "mark everything" shortcut).
fn expr_is_python_float(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Binary { op: BinOp::Div, .. } => true,
        ExprKind::Binary {
            op:
                BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::Mod
                | BinOp::Pow
                | BinOp::FloorDiv,
            left,
            right,
        } => expr_is_python_float(left) || expr_is_python_float(right),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => expr_is_python_float(expr),
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(n) if n == "float" => true,
            // Python `+`/`*` lower to __pyadd__/__pymul__ — float if an operand is.
            ExprKind::Ident(n) if n == "__pyadd__" || n == "__pymul__" => {
                args.iter().any(|a| expr_is_python_float(&a.value))
            }
            ExprKind::Member { object, field, .. } => {
                matches!(&object.kind, ExprKind::Ident(o) if o == "math")
                    && FLOAT_MATH_FNS.contains(&field.as_str())
            }
            _ => false,
        },
        _ => false,
    }
}

/// Wrap `value` in `__py_float_repr__(value)` so it displays Python-float-style.
fn wrap_float_repr(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__py_float_repr__".into()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

/// Tier 2 of float display: like `expr_is_python_float` but also treats a bare
/// variable known (from a prior assignment) to hold a float as a float.
fn expr_is_float_ctx(e: &Expression, floats: &HashMap<String, bool>) -> bool {
    match &e.kind {
        ExprKind::Ident(name) => *floats.get(name).unwrap_or(&false),
        ExprKind::Binary { op: BinOp::Div, .. } => true,
        ExprKind::Binary {
            op: BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Mod | BinOp::Pow,
            left,
            right,
        } => expr_is_float_ctx(left, floats) || expr_is_float_ctx(right, floats),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => expr_is_float_ctx(expr, floats),
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "__pyadd__" || n == "__pymul__") =>
        {
            args.iter().any(|a| expr_is_float_ctx(&a.value, floats))
        }
        _ => expr_is_python_float(e),
    }
}

/// Wrap bare float-variable arguments of a `print(...)` call so they display
/// Python-float-style. (Direct float expressions were already wrapped during
/// `normalize_python_print_args`; here we catch variables tracked in `floats`.)
fn wrap_float_print_vars(e: &mut Expression, floats: &HashMap<String, bool>) {
    if let ExprKind::Call { callee, args, .. } = &mut e.kind {
        if matches!(&callee.kind, ExprKind::Ident(n) if n == "print") {
            // args[0]=sep, args[1]=end, args[2..]=items.
            for a in args.iter_mut().skip(2) {
                if a.name.is_none() && !a.spread {
                    if matches!(&a.value.kind, ExprKind::Ident(name) if *floats.get(name).unwrap_or(&false))
                    {
                        let v = std::mem::replace(&mut a.value, Expression::null());
                        a.value = wrap_float_repr(v);
                    }
                }
            }
        }
    }
}

/// Post-pass: track which local variables hold floats and wrap float-variable
/// `print` arguments. Function bodies get a fresh scope.
fn apply_float_var_repr(stmts: &mut [Statement], floats: &mut HashMap<String, bool>) {
    for stmt in stmts.iter_mut() {
        match &mut stmt.kind {
            StmtKind::Assign { targets, value } => {
                let is_f = expr_is_float_ctx(value, floats);
                if let [t] = targets.as_slice() {
                    if let ExprKind::Ident(name) = &t.kind {
                        floats.insert(name.clone(), is_f);
                    }
                }
            }
            StmtKind::Expr(e) | StmtKind::Return(Some(e)) => wrap_float_print_vars(e, floats),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                apply_float_var_repr(then_body, floats);
                for (_, b) in elifs.iter_mut() {
                    apply_float_var_repr(b, floats);
                }
                if let Some(b) = else_body {
                    apply_float_var_repr(b, floats);
                }
            }
            StmtKind::While { body, .. } | StmtKind::ForIn { body, .. } => {
                apply_float_var_repr(body, floats)
            }
            StmtKind::For { body, .. } => apply_float_var_repr(body, floats),
            StmtKind::Block(b) => apply_float_var_repr(b, floats),
            StmtKind::FunctionDecl { body, .. } => {
                let mut inner = HashMap::new();
                apply_float_var_repr(body, &mut inner);
            }
            _ => {}
        }
    }
}

fn normalize_python_print_args(raw: Vec<Argument>) -> Vec<Argument> {
    let mut sep = Argument::positional(Expression::string(" "));
    let mut end = Argument::positional(Expression::string("\n"));
    let mut items = Vec::new();
    for a in raw {
        match a.name.as_deref() {
            Some("sep") => sep = Argument::positional(a.value),
            Some("end") => end = Argument::positional(a.value),
            Some("file") | Some("flush") => {}
            _ => {
                // Display statically-known bytes as `b'…'`, floats as `4.0`.
                if expr_is_python_bytes(&a.value) {
                    items.push(Argument::positional(wrap_bytes_repr(a.value)));
                } else if expr_is_python_float(&a.value) {
                    items.push(Argument::positional(wrap_float_repr(a.value)));
                } else {
                    items.push(a);
                }
            }
        }
    }
    let mut out = Vec::with_capacity(items.len() + 2);
    out.push(sep);
    out.push(end);
    out.extend(items);
    out
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            // The `*` / `**` prefixes are silent literals in the grammar, so
            // check the call_arg's own text rather than a child token.
            let arg_text = p.as_str().trim_start();
            let is_starstar = arg_text.starts_with("**");
            let is_star = !is_starstar && arg_text.starts_with('*');
            let mut ci: Vec<Pair<Rule>> = p.into_inner().collect();
            if ci.is_empty() {
                continue;
            }

            if is_starstar {
                // **kwargs — spread of a mapping into keyword arguments.
                let val = walk_expression(ci.pop().unwrap())?;
                let val = match val.kind {
                    ExprKind::Spread(inner) => *inner,
                    _ => val,
                };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            } else if is_star {
                // *args — positional spread expansion.
                let val = walk_expression(ci.pop().unwrap())?;
                let val = match val.kind {
                    ExprKind::Spread(inner) => *inner,
                    _ => val,
                };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
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
                // `*args` parses as a `star_expr` → `ExprKind::Spread`; unwrap it
                // and flag the argument as spread so the call expands it.
                let val = walk_expression(ci.remove(0))?;
                if let ExprKind::Spread(inner) = val.kind {
                    args.push(Argument {
                        value: *inner,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                } else {
                    args.push(Argument::positional(val));
                }
            }
        }
    }
    Ok(args)
}

// ── Subscript ───────────────────────────────────────────────────────────────

/// Wrap a scalar subscript index in the from-end offset normalizer
/// `__py_from_end__` so `a[-1]` reads one-from-the-end (like C#'s `arr[^N]`).
/// Skips keys that can never be a from-end offset: string literals (dict keys),
/// non-negative integer literals, and slices/ranges (the slice path already
/// offsets from the end). The normalizer is a runtime no-op unless the index is
/// a negative number on a sequence, so dict lookups stay direct.
fn python_index_operand(object: &Expression, index: Expression) -> Expression {
    match &index.kind {
        ExprKind::Lit(Literal::Str(_)) => return index,
        ExprKind::Lit(Literal::Int(n)) if *n >= 0 => return index,
        ExprKind::Slice { .. } | ExprKind::Range { .. } => return index,
        _ => {}
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__py_from_end__")),
        args: vec![
            Argument::positional(object.clone()),
            Argument::positional(index),
        ],
        optional: false,
    })
}

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

    // Normal list. `*x` elements walk to `ExprKind::Spread(x)`; unwrap them and
    // set the `spread` flag so `[*a, *b]` flattens instead of nesting.
    let elements = inner
        .into_iter()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            if let ExprKind::Spread(inner) = val.kind {
                Ok(ArrayElement {
                    key: None,
                    value: *inner,
                    spread: true,
                    by_ref: false,
                })
            } else {
                Ok(ArrayElement {
                    key: None,
                    value: val,
                    spread: false,
                    by_ref: false,
                })
            }
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
    // Python quirk: empty `{}` is an empty DICT, not a set (`set()` is the
    // empty set). Without this, `{}` became `ExprKind::Set([])` and `d[k]=v`
    // never created enumerable object properties.
    let mut is_dict = inner.is_empty();
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

    // Wrap string literals in [...s] so comprehensions iterate chars
    if matches!(iter.kind, ExprKind::Lit(Literal::Str(_))) {
        iter = Expression::new(ExprKind::Array(vec![ArrayElement {
            key: None,
            spread: true,
            by_ref: false,
            value: iter,
        }]));
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

/// Convert an assignment-target expression element into a destructuring
/// pattern element, handling `*rest` (`ExprKind::Spread` → `Rest`) and nested
/// tuple/list targets (`(a, (b, c))` → nested `Array` pattern) — not just bare
/// identifiers.
fn expr_to_array_pattern_elem(e: &Expression) -> ArrayPatternElem {
    match &e.kind {
        ExprKind::Ident(name) => {
            ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
        }
        ExprKind::Spread(inner) => match &inner.kind {
            ExprKind::Ident(name) => ArrayPatternElem::Rest(name.clone()),
            _ => ArrayPatternElem::Hole,
        },
        ExprKind::Tuple(elems) => {
            let nested = elems.iter().map(expr_to_array_pattern_elem).collect();
            ArrayPatternElem::Pattern(BindingPattern::Array(nested), None)
        }
        ExprKind::Array(elems) => {
            let nested = elems
                .iter()
                .map(|ae| expr_to_array_pattern_elem(&ae.value))
                .collect();
            ArrayPatternElem::Pattern(BindingPattern::Array(nested), None)
        }
        _ => ArrayPatternElem::Hole,
    }
}

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
    wrap_bytes(Expression::new(ExprKind::Array(elements))).kind
}

/// Wrap an int-array expression in the `__py_bytes__` marker so downstream
/// static analysis (`expr_is_python_bytes`) can distinguish bytes from a list.
/// The marker is an identity passthrough at runtime.
fn wrap_bytes(array: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__py_bytes__".into()))),
        args: vec![Argument::positional(array)],
        optional: false,
    })
}

/// True when `e` is statically known to evaluate to `bytes`: a `b'…'` literal
/// (already wrapped in `__py_bytes__`), a `bytes(...)` call, `str.encode()`, or
/// a `+`/`*`/slice built from bytes.
fn expr_is_python_bytes(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(n) if n == "__py_bytes__" || n == "bytes" => true,
            // `+`/`*` lower to __pyadd__/__pymul__ — bytes if an operand is.
            ExprKind::Ident(n) if n == "__pyadd__" || n == "__pymul__" => {
                args.iter().any(|a| expr_is_python_bytes(&a.value))
            }
            ExprKind::Member { field, .. } if field == "encode" => true,
            _ => false,
        },
        // slice / index of bytes stays bytes for slices; a plain index is an int
        // (handled by the caller, which only wraps whole-value repr contexts).
        ExprKind::Index { object, .. } => expr_is_python_bytes(object),
        _ => false,
    }
}

/// Wrap a bytes-valued expression in a `__vybe_bytes_repr(...)` call so it
/// displays as `b'…'`.
fn wrap_bytes_repr(e: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__vybe_bytes_repr".into()))),
        args: vec![Argument::positional(e)],
        optional: false,
    })
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
        "is" => BinOp::Eq,
        _ if s.contains("not") && s.contains("in") => BinOp::NotIn,
        _ if s.contains("is") && s.contains("not") => BinOp::NotEq,
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
