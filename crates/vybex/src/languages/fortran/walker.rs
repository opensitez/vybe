//! Fortran walker — pest `Pair<Rule>` → `vybex::ast::Module`.

use pest::Parser;
use pest::iterators::Pair;
use crate::ast::*;
use super::{FortranParser, Rule};

/// Convert legacy Hollerith literals like `4HTEST` to standard string literals
/// `"TEST"`. Pest can't match a runtime-determined character count, so we
/// preprocess the source. We respect string literals (don't rewrite inside
/// `'...'` or `"..."`) and comments (anything after `!` to end of line).
fn rewrite_hollerith(src: &str) -> String {
    let mut out = String::with_capacity(src.len());
    let bytes = src.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        // Skip line comments.
        if b == b'!' {
            while i < bytes.len() && bytes[i] != b'\n' {
                out.push(bytes[i] as char);
                i += 1;
            }
            continue;
        }
        // Skip string literals.
        if b == b'\'' || b == b'"' {
            let quote = b;
            out.push(b as char);
            i += 1;
            while i < bytes.len() {
                let c = bytes[i];
                out.push(c as char);
                i += 1;
                if c == quote {
                    // Fortran doubled-quote escape.
                    if i < bytes.len() && bytes[i] == quote {
                        out.push(quote as char);
                        i += 1;
                        continue;
                    }
                    break;
                }
            }
            continue;
        }
        // Try Hollerith: must be at start of token (preceded by non-alnum).
        if b.is_ascii_digit() {
            let prev_is_word = i > 0 && {
                let p = bytes[i - 1];
                p.is_ascii_alphanumeric() || p == b'_' || p == b'.'
            };
            if !prev_is_word {
                let mut j = i;
                while j < bytes.len() && bytes[j].is_ascii_digit() { j += 1; }
                if j < bytes.len() && (bytes[j] == b'H' || bytes[j] == b'h') {
                    if let Ok(count) = std::str::from_utf8(&bytes[i..j]).unwrap().parse::<usize>() {
                        let after_h = j + 1;
                        let end = (after_h + count).min(bytes.len());
                        if end - after_h == count {
                            let text = &src[after_h..end];
                            // Emit as quoted string with escaped quotes.
                            out.push('"');
                            for ch in text.chars() {
                                if ch == '"' { out.push('\\'); }
                                out.push(ch);
                            }
                            out.push('"');
                            i = end;
                            continue;
                        }
                    }
                }
            }
        }
        out.push(b as char);
        i += 1;
    }
    out
}

pub fn parse(source: &str) -> Result<Module, String> {
    let preprocessed = rewrite_hollerith(source);
    let source = preprocessed.as_str();
    let mut pairs = FortranParser::parse(Rule::program, source)
        .map_err(|e| format!("Fortran parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut name = String::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI | Rule::NEWLINE => {}
            Rule::statement_line => {
                for inner in pair.into_inner().filter(|p| meaningful(p)) {
                    walk_top(inner, &mut name, &mut body, &mut imports)?;
                }
            }
            _ => {}
        }
    }

    Ok(Module { name, language: Lang::Fortran, body, imports })
}

fn walk_top(pair: Pair<Rule>, name: &mut String, body: &mut Vec<Statement>, imports: &mut Vec<Import>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::program_unit => {
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => { if name.is_empty() { *name = p.as_str().to_string(); } }
                    Rule::statement_line => {
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            walk_top(s, name, body, imports)?;
                        }
                    }
                    _ => { if let Some(st) = walk_stmt(p)? { body.push(st); } }
                }
            }
        }
        Rule::module_unit => {
            let mut mname = String::new();
            let mut members = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => { if mname.is_empty() { mname = p.as_str().to_string(); } }
                    Rule::statement_line => {
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? {
                                members.push(to_class_member(st));
                            }
                        }
                    }
                    _ => { if let Some(st) = walk_stmt(p)? { members.push(to_class_member(st)); } }
                }
            }
            body.push(Statement::new(StmtKind::ModuleDecl { name: mname, members, visibility: Visibility::Public }));
        }
        Rule::use_statement => {
            let mut parts = pair.into_inner().filter(|p| meaningful(p));
            let mname = parts.next().map(|p| p.as_str().to_string()).unwrap_or_default();
            let mut names = Vec::new();
            for p in parts {
                if p.as_rule() == Rule::use_name_list {
                    for np in p.into_inner() {
                        if np.as_rule() == Rule::use_name {
                            let mut ni = np.into_inner();
                            let n = ni.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                            let a = ni.next().map(|p| p.as_str().to_string());
                            names.push(ImportName { name: n, alias: a });
                        }
                    }
                }
            }
            if names.is_empty() {
                imports.push(Import { kind: ImportKind::Simple { path: mname, alias: None }, span: Span::default() });
            } else {
                imports.push(Import { kind: ImportKind::Named { path: mname, names, level: 0 }, span: Span::default() });
            }
        }
        _ => { if let Some(st) = walk_stmt(pair)? { body.push(st); } }
    }
    Ok(())
}

fn walk_stmt(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::implicit_statement | Rule::contains_statement | Rule::NEWLINE => Ok(None),
        Rule::var_declaration => walk_var_decl(pair).map(Some),
        Rule::assignment_statement => walk_assign(pair).map(Some),
        Rule::if_statement => walk_if(pair).map(Some),
        Rule::do_statement => walk_do(pair).map(Some),
        Rule::do_while_statement => walk_do_while(pair).map(Some),
        Rule::select_case_statement => walk_select(pair).map(Some),
        Rule::print_statement | Rule::write_statement => walk_print(pair).map(Some),
        Rule::read_statement => Ok(Some(Statement::new(StmtKind::Empty))),
        Rule::call_statement => walk_call(pair).map(Some),
        Rule::subroutine_decl => walk_sub(pair).map(Some),
        Rule::function_decl => walk_func(pair).map(Some),
        Rule::type_decl => walk_type(pair).map(Some),
        Rule::interface_decl | Rule::allocate_statement | Rule::deallocate_statement => Ok(Some(Statement::new(StmtKind::Empty))),
        Rule::return_statement => {
            let e = pair.into_inner().filter(|p| meaningful(p)).next().map(walk_expr).transpose()?;
            Ok(Some(Statement::new(StmtKind::Return(e))))
        }
        Rule::cycle_statement => Ok(Some(Statement::new(StmtKind::Continue(ContinueTarget::Implicit)))),
        Rule::exit_statement => Ok(Some(Statement::new(StmtKind::Break(BreakTarget::Implicit)))),
        Rule::stop_statement => Ok(Some(Statement::new(StmtKind::Return(None)))),
        Rule::expression_statement => {
            let e = walk_expr(pair.into_inner().next().ok_or("empty expr")?)?;
            Ok(Some(Statement::new(StmtKind::Expr(e))))
        }
        Rule::statement_line => {
            let mut stmts = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if let Some(s) = walk_stmt(p)? { stmts.push(s); }
            }
            match stmts.len() {
                0 => Ok(None),
                1 => Ok(stmts.into_iter().next()),
                _ => Ok(Some(Statement::new(StmtKind::Block(stmts)))),
            }
        }
        Rule::use_statement => Ok(None),
        Rule::program_unit | Rule::module_unit => {
            let mut body = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if p.as_rule() == Rule::statement_line {
                    for s in p.into_inner().filter(|p| meaningful(p)) {
                        if let Some(st) = walk_stmt(s)? { body.push(st); }
                    }
                } else if p.as_rule() != Rule::identifier {
                    if let Some(st) = walk_stmt(p)? { body.push(st); }
                }
            }
            Ok(Some(Statement::new(StmtKind::Block(body))))
        }
        _ => Ok(None),
    }
}

fn walk_body<'a>(pairs: impl Iterator<Item = Pair<'a, Rule>>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pairs {
        match p.as_rule() {
            Rule::statement_line => {
                for s in p.into_inner().filter(|p| meaningful(p)) {
                    if let Some(st) = walk_stmt(s)? { body.push(st); }
                }
            }
            Rule::identifier => {}
            _ => { if let Some(st) = walk_stmt(p)? { body.push(st); } }
        }
    }
    Ok(body)
}

/// Extract the declared length N from a `character(len=N)` or
/// `character*N` type hint, lowercased and whitespace-stripped.
/// Returns `None` for `character` (no length), `character(len=*)`,
/// `character(len=:)`, or non-character types.
fn parse_character_len(type_hint: &str) -> Option<i64> {
    let s: String = type_hint.chars().filter(|c| !c.is_whitespace()).collect();
    let lower = s.to_ascii_lowercase();
    if !lower.starts_with("character") {
        return None;
    }
    // `character*N` form
    if let Some(rest) = lower.strip_prefix("character*") {
        let num: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
        if !num.is_empty() {
            return num.parse().ok();
        }
    }
    // `character(len=N)` or `character(N)` form. Take what's inside the
    // outermost parens and look for a numeric length.
    let lp = lower.find('(')?;
    let rp = lower.rfind(')')?;
    if rp <= lp {
        return None;
    }
    let inside = &lower[lp + 1..rp];
    // Common forms: `len=N`, `n,kind=k`, `N`, `len=*`, `len=:`.
    let first_clause = inside.split(',').next().unwrap_or(inside);
    let val = if let Some(eq) = first_clause.split_once('=') {
        eq.1
    } else {
        first_clause
    };
    if val == "*" || val == ":" {
        return None;
    }
    val.parse().ok()
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner();
    let type_hint = inner.next().map(|p| p.as_str().trim().to_string());
    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator_list {
            for d in p.into_inner() {
                if d.as_rule() == Rule::var_declarator {
                    let mut di = d.into_inner().filter(|p| meaningful(p));
                    let nm = di.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                    let mut init = None;
                    let mut dim_size: Option<Expression> = None;
                    for pp in di {
                        match pp.as_rule() {
                            Rule::dimension_spec_list => {
                                // Compute total size: product of dimensions.
                                // Each dimension_spec is either lower:upper, expression (size), or `:` (deferred).
                                let mut size_expr: Option<Expression> = None;
                                for spec in pp.into_inner().filter(|p| meaningful(p)) {
                                    if spec.as_rule() != Rule::dimension_spec { continue; }
                                    let exprs: Vec<Pair<Rule>> = spec.clone().into_inner()
                                        .filter(|p| meaningful(p)).collect();
                                    let this_size = if exprs.len() == 1 {
                                        // bare expression — that's the upper bound, size = it
                                        walk_expr(exprs.into_iter().next().unwrap())?
                                    } else if exprs.len() == 2 {
                                        // lower:upper => upper - lower + 1
                                        let lo = walk_expr(exprs[0].clone())?;
                                        let hi = walk_expr(exprs[1].clone())?;
                                        let sub = Expression::new(ExprKind::Binary {
                                            op: BinOp::Sub,
                                            left: Box::new(hi),
                                            right: Box::new(lo),
                                        });
                                        Expression::new(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(sub),
                                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                                        })
                                    } else {
                                        // deferred or unrecognized — skip
                                        continue;
                                    };
                                    size_expr = Some(match size_expr.take() {
                                        Some(prev) => Expression::new(ExprKind::Binary {
                                            op: BinOp::Mul,
                                            left: Box::new(prev),
                                            right: Box::new(this_size),
                                        }),
                                        None => this_size,
                                    });
                                }
                                dim_size = size_expr;
                            }
                            Rule::codimension_spec_list => { /* PGAS — model as scalar/array */ }
                            _ => {
                                init = Some(walk_expr(pp)?);
                            }
                        }
                    }
                    // If declared with dimensions but no explicit init, synthesize Array(sz, 0).
                    // The compiler intercepts this 2-arg pattern (used by COBOL OCCURS) and
                    // emits ecma:array.newWithLength + fill via compiler_common.
                    if init.is_none() {
                        if let Some(sz) = dim_size {
                            init = Some(Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Ident("Array".into()))),
                                args: vec![
                                    Argument::positional(sz),
                                    Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                                ],
                                optional: false,
                            }));
                        }
                    }
                    // Fortran `character(len=N) :: s = 'literal'` — pad the literal
                    // with trailing blanks so `len(s)` returns the declared length.
                    // Pure JS-shape rewrite: wrap the init in `s.padEnd(N, ' ')`.
                    if let Some(ref t) = type_hint {
                        if let Some(declared_len) = parse_character_len(t) {
                            if let Some(ref existing) = init {
                                let padded = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: Box::new(existing.clone()),
                                        field: "padEnd".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![
                                        Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(declared_len)))),
                                        Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(" ".into())))),
                                    ],
                                    optional: false,
                                });
                                init = Some(padded);
                            } else {
                                // No init — synthesize a string of N spaces so len(s) returns N.
                                let spaces: String = " ".repeat(declared_len as usize);
                                init = Some(Expression::new(ExprKind::Lit(Literal::Str(spaces))));
                            }
                        }
                    }
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(nm),
                        type_hint: type_hint.clone(),
                        init,
                        array_bounds: None,
                        with_events: false,
                    });
                }
            }
        }
    }
    Ok(Statement::new(StmtKind::VarDecl { declarations, kind: VarDeclKind::Dim }))
}

fn walk_assign(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let value = walk_expr(parts.pop().ok_or("missing rhs")?)?;
    let mut target = Expression::new(ExprKind::Ident(parts[0].as_str().to_string()));
    for p in &parts[1..] {
        if p.as_rule() == Rule::member_or_index {
            for m in p.clone().into_inner() {
                if m.as_rule() == Rule::identifier {
                    target = Expression::new(ExprKind::Member {
                        object: Box::new(target), field: m.as_str().to_string(), null_safe: false,
                    });
                } else {
                    let idx = walk_expr(m)?;
                    target = Expression::new(ExprKind::Index { object: Box::new(target), index: Box::new(idx), null_safe: false });
                }
            }
        }
    }
    Ok(Statement::new(StmtKind::Assign { targets: vec![target], value }))
}

fn walk_if(pair: Pair<Rule>) -> Result<Statement, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // Find the condition expression — first non-keyword child
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;
    for p in inner {
        match p.as_rule() {
            Rule::expression | Rule::logical_or | Rule::logical_and | Rule::logical_not
            | Rule::comparison | Rule::addition | Rule::multiplication | Rule::power
            | Rule::concat | Rule::unary | Rule::primary_expr => {
                if cond.is_none() {
                    cond = Some(walk_expr(p)?);
                }
            }
            Rule::statement_line => {
                for s in p.into_inner().filter(|p| meaningful(p)) {
                    if let Some(st) = walk_stmt(s)? { then_body.push(st); }
                }
            }
            Rule::elseif_clause => {
                let ei: Vec<Pair<Rule>> = p.into_inner().filter(|p| meaningful(p)).collect();
                let mut ec = None;
                let mut eb = Vec::new();
                for e in ei {
                    if is_expr_rule(e.as_rule()) && ec.is_none() {
                        ec = Some(walk_expr(e)?);
                    } else if e.as_rule() == Rule::statement_line {
                        for s in e.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { eb.push(st); }
                        }
                    }
                }
                if let Some(c) = ec {
                    elifs.push((c, eb));
                }
            }
            Rule::else_clause => {
                let mut eb = Vec::new();
                for e in p.into_inner().filter(|p| meaningful(p)) {
                    if e.as_rule() == Rule::statement_line {
                        for s in e.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { eb.push(st); }
                        }
                    }
                }
                else_body = Some(eb);
            }
            // Single-line if body (e.g., print_statement, assignment_statement)
            Rule::print_statement | Rule::write_statement | Rule::call_statement
            | Rule::assignment_statement | Rule::return_statement | Rule::cycle_statement
            | Rule::exit_statement | Rule::stop_statement | Rule::expression_statement => {
                if let Some(st) = walk_stmt(p)? { then_body.push(st); }
            }
            _ => {} // skip keywords (kw_if, kw_then, kw_end, etc.)
        }
    }
    Ok(Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body, elifs, else_body
    }))
}

fn walk_do(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // Collect: identifier, expression, expression [, expression], statement_line*
    let mut var = String::new();
    let mut exprs = Vec::new();
    let mut body_parts = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::identifier if var.is_empty() => { var = p.as_str().to_string(); }
            Rule::statement_line => { body_parts.push(p); }
            _ if is_expr_rule(p.as_rule()) => { exprs.push(p); }
            Rule::identifier => {} // end do name
            _ => {} // skip kw_do, kw_end etc.
        }
    }
    let start = if !exprs.is_empty() { walk_expr(exprs.remove(0))? } else { Expression::new(ExprKind::Lit(Literal::Int(0))) };
    let end_e = if !exprs.is_empty() { walk_expr(exprs.remove(0))? } else { Expression::new(ExprKind::Lit(Literal::Int(0))) };
    let step_expr = if !exprs.is_empty() { Some(walk_expr(exprs.remove(0))?) } else { None };
    let body = walk_body(body_parts.into_iter())?;
    let init = Some(Box::new(Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Ident(var.clone()))],
        value: start,
    })));
    let cond = Some(Expression::new(ExprKind::Binary {
        left: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
        op: BinOp::LtEq,
        right: Box::new(end_e),
    }));
    let sv = step_expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
    // i = i + step as an Assign expression
    let update = Some(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
        value: Box::new(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Ident(var))),
            op: BinOp::Add,
            right: Box::new(sv),
        })),
    }));
    Ok(Statement::new(StmtKind::For { init, cond, update, body }))
}

fn walk_do_while(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut cond = None;
    let mut body_parts = Vec::new();
    for p in parts {
        if is_expr_rule(p.as_rule()) && cond.is_none() {
            cond = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::statement_line {
            body_parts.push(p);
        }
        // skip kw_do, kw_while, kw_end
    }
    // If condition not found, emit "false" so the loop immediately exits (never infinite)
    let cond = cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(false))));
    let body = walk_body(body_parts.into_iter())?;
    Ok(Statement::new(StmtKind::While { cond, body, else_body: None }))
}

fn walk_select(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut expr = None;
    let mut case_pairs = Vec::new();
    for p in parts {
        if is_expr_rule(p.as_rule()) && expr.is_none() {
            expr = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::case_block {
            case_pairs.push(p);
        }
    }
    let expr = expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
    let mut cases = Vec::new();
    for p in case_pairs {
        {
            let mut conds = Vec::new();
            let mut cbody = Vec::new();
            let mut is_default = false;
            for c in p.into_inner().filter(|p| meaningful(p)) {
                match c.as_rule() {
                    Rule::case_value_list => {
                        for cv in c.into_inner() {
                            if cv.as_rule() == Rule::case_value {
                                if let Some(first) = cv.into_inner().next() {
                                    conds.push(CaseCondition::Value(walk_expr(first)?));
                                }
                            }
                        }
                    }
                    Rule::statement_line => {
                        for s in c.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { cbody.push(st); }
                        }
                    }
                    _ => { if c.as_str().to_lowercase().contains("default") { is_default = true; } }
                }
            }
            if is_default { conds.clear(); }
            cases.push(SwitchCase { conditions: conds, body: cbody });
        }
    }
    Ok(Statement::new(StmtKind::Switch { expr, cases, default: None }))
}

fn walk_print(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut args = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        if is_expr_rule(p.as_rule()) {
            args.push(Argument::positional(walk_expr(p)?));
        }
    }
    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__print".into()))),
        args, optional: false,
    }))))
}

fn walk_call(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner().filter(|p| meaningful(p));
    let nm = inner.next().ok_or("missing call name")?.as_str().to_string();
    let mut args = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::argument_list {
            for a in p.into_inner() {
                if a.as_rule() == Rule::argument {
                    let mut ai: Vec<Pair<Rule>> = a.into_inner().filter(|p| meaningful(p)).collect();
                    if ai.len() >= 1 {
                        let v = walk_expr(ai.pop().unwrap())?;
                        let n = if !ai.is_empty() { Some(ai[0].as_str().to_string()) } else { None };
                        args.push(Argument { name: n, value: v, by_ref: false, spread: false });
                    }
                }
            }
        }
    }
    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(nm))),
        args, optional: false,
    }))))
}

fn walk_sub(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        if p.as_rule() == Rule::identifier && nm.is_empty() {
            nm = p.as_str().to_string();
        } else if p.as_rule() == Rule::param_list {
            for pp in p.into_inner() {
                if pp.as_rule() == Rule::identifier {
                    params.push(Param { name: pp.as_str().to_string(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false });
                }
            }
        } else { rest.push(p); }
    }
    let body = walk_body(rest.into_iter())?;
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm, params, return_type: None, body, modifiers: Modifiers::default(),
        handles: vec![], is_async: false, is_generator: false, is_sub: true,
    }))
}

fn walk_func(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rt = None;
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::type_spec => { rt = Some(p.as_str().trim().to_string()); }
            Rule::identifier => { if nm.is_empty() { nm = p.as_str().to_string(); } }
            Rule::param_list => {
                for pp in p.into_inner() {
                    if pp.as_rule() == Rule::identifier {
                        params.push(Param { name: pp.as_str().to_string(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false });
                    }
                }
            }
            Rule::result_clause => {}
            _ => { rest.push(p); }
        }
    }
    let body = walk_body(rest.into_iter())?;
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm, params, return_type: rt, body, modifiers: Modifiers::default(),
        handles: vec![], is_async: false, is_generator: false, is_sub: false,
    }))
}

fn walk_type(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut nm = String::new();
    let mut members = Vec::new();
    let mut parents = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::identifier => { if nm.is_empty() { nm = p.as_str().to_string(); } }
            Rule::type_attribute => {
                for a in p.into_inner() { if a.as_rule() == Rule::identifier { parents.push(a.as_str().to_string()); } }
            }
            Rule::type_member => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::var_declaration {
                        let decl = walk_var_decl(m)?;
                        if let StmtKind::VarDecl { declarations, .. } = &decl.kind {
                            for d in declarations {
                                if let BindingPattern::Ident(fname) = &d.pattern {
                                    members.push(ClassMember::Field {
                                        name: fname.clone(), type_hint: d.type_hint.clone(), init: d.init.clone(),
                                        modifiers: Modifiers::default(), with_events: false, array_bounds: None,
                                    });
                                }
                            }
                        }
                    }
                }
            }
            Rule::type_bound_procedure => {
                let mn = p.into_inner().filter(|p| meaningful(p))
                    .find(|p| p.as_rule() == Rule::identifier)
                    .map(|p| p.as_str().to_string()).unwrap_or_default();
                if !mn.is_empty() {
                    members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                        name: mn, params: vec![], return_type: None, body: vec![],
                        modifiers: Modifiers::default(), handles: vec![], is_async: false, is_generator: false, is_sub: true,
                    }))));
                }
            }
            _ => {}
        }
    }
    Ok(Statement::new(StmtKind::ClassDecl {
        name: nm, parents, interfaces: vec![], members, modifiers: ClassModifiers::default(),
    }))
}

// ── Expressions ────────────────────────────────────────────────────────────

fn walk_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::expression | Rule::logical_or | Rule::logical_and | Rule::logical_not
        | Rule::comparison | Rule::addition | Rule::multiplication | Rule::power | Rule::concat
        | Rule::unary => walk_binop(pair),
        Rule::primary_expr => {
            // primary_atom followed by zero or more postfix_op (member, call, coindex).
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let atom = inner.next().ok_or("empty primary")?;
            let mut expr = walk_expr(atom)?;
            for op in inner {
                if op.as_rule() != Rule::postfix_op { continue; }
                let mut op_inner = op.clone().into_inner().filter(|p| meaningful(p));
                let first = op_inner.next();
                match first {
                    Some(p) if p.as_rule() == Rule::identifier => {
                        // %field — member access
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                    Some(p) if p.as_rule() == Rule::argument_list => {
                        // (args) — call or index. Treat as Call by default;
                        // codegen for arrays sees Call(arr, [idx]) and emits index access.
                        let mut args = Vec::new();
                        for a in p.into_inner() {
                            if a.as_rule() == Rule::argument {
                                let v = a.into_inner().filter(|q| meaningful(q)).last()
                                    .ok_or("empty arg")?;
                                args.push(Argument::positional(walk_expr(v)?));
                            }
                        }
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args,
                            optional: false,
                        });
                    }
                    Some(_) => { /* ignore */ }
                    None => {
                        // Empty `()` — call with no args
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args: Vec::new(),
                            optional: false,
                        });
                    }
                }
            }
            Ok(expr)
        }
        Rule::primary_atom => walk_expr(pair.into_inner().next().ok_or("empty atom")?),
        Rule::complex_literal => {
            // `(re, im)` complex constant — synthesise as a call to a
            // compiler-known `cmplx(re, im)` so the rest of the
            // pipeline treats it like the standard intrinsic. Walker
            // doesn't carry a Complex literal kind — the AST stays
            // language-neutral.
            let mut parts = pair.into_inner().filter(|p| meaningful(p));
            let re = walk_expr(parts.next().ok_or("complex: missing real")?)?;
            let im = walk_expr(parts.next().ok_or("complex: missing imag")?)?;
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("cmplx".to_string()))),
                args: vec![Argument::positional(re), Argument::positional(im)],
                optional: false,
            }));
        }
        Rule::array_constructor => {
            // `[a, b, c]` / `(/ a, b, c /)` / implied-do — walker
            // builds a flat `ExprKind::Array`. Implied-do `(i, i=1,n)`
            // is collapsed to its trip-count expression for now (the
            // proper expansion is a comprehension, pending a
            // walker-level rewrite to a generated loop).
            let mut elems: Vec<crate::ast::ArrayElement> = Vec::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::array_constructor_body) {
                    for v in p.into_inner() {
                        if matches!(v.as_rule(), Rule::array_constructor_value) {
                            // Walk the first expression child as the
                            // element value (implied-do collapses to
                            // its body expression for the MVP).
                            if let Some(inner) = v.into_inner()
                                .filter(|q| meaningful(q))
                                .find(|q| is_expr_rule(q.as_rule())
                                    || matches!(q.as_rule(), Rule::expression))
                            {
                                elems.push(crate::ast::ArrayElement {
                                    key: None,
                                    value: walk_expr(inner)?,
                                    spread: false,
                                    by_ref: false,
                                });
                            }
                        }
                    }
                }
            }
            Ok(Expression::new(ExprKind::Array(elems)))
        }
        Rule::literal => walk_expr(pair.into_inner().next().ok_or("empty literal")?),
        Rule::logical_literal => {
            Ok(Expression::new(ExprKind::Lit(Literal::Bool(pair.as_str().to_lowercase().contains("true")))))
        }
        Rule::number_literal => {
            let s = pair.as_str().trim();
            let clean = s.split('_').next().unwrap_or(s);
            if clean.contains('.') || clean.to_lowercase().contains('e') || clean.to_lowercase().contains('d') {
                let n: f64 = clean.replace('d', "e").replace('D', "E").parse().unwrap_or(0.0);
                Ok(Expression::new(ExprKind::Lit(Literal::Float(n))))
            } else {
                let n: i64 = clean.parse().unwrap_or(0);
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            }
        }
        Rule::string_literal => {
            let s = pair.as_str();
            let inner = &s[1..s.len()-1];
            Ok(Expression::new(ExprKind::Lit(Literal::Str(inner.replace("''", "'").replace("\"\"", "\"")))))
        }
        Rule::boz_literal => {
            // `b'..'` / `o'..'` / `z'..'` — bit / octal / hex literal.
            let s = pair.as_str();
            let prefix = s.chars().next().unwrap_or('z').to_ascii_lowercase();
            let body = &s[1..];
            let trimmed = body.trim_matches(|c: char| c == '\'' || c == '"');
            let radix = match prefix { 'b' => 2, 'o' => 8, _ => 16 };
            let n = i64::from_str_radix(trimmed, radix).unwrap_or(0);
            Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
        }
        Rule::identifier => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::function_call_or_subscript => {
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let nm = inner.next().ok_or("missing fn")?.as_str().to_string();
            let mut args = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for a in p.into_inner() {
                        if a.as_rule() == Rule::argument {
                            let v = a.into_inner().filter(|p| meaningful(p)).last().ok_or("empty arg")?;
                            args.push(Argument::positional(walk_expr(v)?));
                        }
                    }
                }
            }
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(nm))), args, optional: false,
            }))
        }
        Rule::argument => walk_expr(pair.into_inner().filter(|p| meaningful(p)).last().ok_or("empty arg")?),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn walk_binop(pair: Pair<Rule>) -> Result<Expression, String> {
    let rule = pair.as_rule();
    // Capture the source slice BEFORE `into_inner()` consumes the
    // pair — `unary`'s inline `-`/`+` literals don't survive as
    // child pairs, so we have to read them from the rule's own span.
    let raw = pair.as_str().to_string();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    // Unary not — `not_op ~ comparison`. inner has [not_op, operand].
    if rule == Rule::logical_not && inner.len() == 2 {
        return Ok(Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(walk_expr(inner.remove(1))?) }));
    }
    // Unary minus/plus — `unary = "-" primary | "+" primary | primary`.
    // Pest doesn't emit inline-literal sign tokens as child pairs,
    // so `inner` is always [primary_expr]. Recover the sign from the
    // rule's leading source character.
    if rule == Rule::unary {
        let trimmed = raw.trim_start();
        if trimmed.starts_with('-') {
            let operand = walk_expr(inner.remove(0))?;
            return Ok(Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg, expr: Box::new(operand),
            }));
        }
        // `+` is a no-op; the `primary` form is the bare value.
        return walk_expr(inner.remove(0));
    }
    // `power = { concat ~ ("**" ~ concat)? }` — inline `**` literal,
    // so inner has either [base] or [base, exponent]. Apply Pow when 2.
    if rule == Rule::power && inner.len() == 2 {
        let base = walk_expr(inner.remove(0))?;
        let exp  = walk_expr(inner.remove(0))?;
        return Ok(Expression::new(ExprKind::Binary {
            left: Box::new(base),
            op: BinOp::Pow,
            right: Box::new(exp),
        }));
    }
    // `concat = { unary ~ ("//" ~ unary)* }` — inline `//` literal,
    // so inner has [u1, u2, u3, ...] without operator pairs. Fold
    // left as Concat (BinOp::Concat).
    if rule == Rule::concat && inner.len() >= 2 {
        let mut result = walk_expr(inner.remove(0))?;
        for next in inner.into_iter() {
            let right = walk_expr(next)?;
            result = Expression::new(ExprKind::Binary {
                left: Box::new(result),
                op: BinOp::Concat,
                right: Box::new(right),
            });
        }
        return Ok(result);
    }
    if inner.len() == 1 { return walk_expr(inner.remove(0)); }
    if inner.len() >= 3 {
        let mut result = walk_expr(inner.remove(0))?;
        let mut i = 0;
        while i + 1 < inner.len() {
            let op = to_binop(&inner[i]);
            let right = walk_expr(inner[i + 1].clone())?;
            result = Expression::new(ExprKind::Binary { left: Box::new(result), op, right: Box::new(right) });
            i += 2;
        }
        return Ok(result);
    }
    if inner.is_empty() { return Ok(Expression::new(ExprKind::Lit(Literal::Null))); }
    walk_expr(inner.remove(0))
}

fn to_binop(pair: &Pair<Rule>) -> BinOp {
    // Pest's `add_op` / `mul_op` / etc. spans can include trailing
    // whitespace before the next operand (the grammar rule is
    // `add_op ~ multiplication`; pest's WHITESPACE-implicit consume
    // can land inside the op's span). Trim before matching so
    // `"- "` still maps to `Sub`, not falling through to the Add
    // default.
    match pair.as_str().to_lowercase().trim() {
        "+" => BinOp::Add, "-" => BinOp::Sub, "*" => BinOp::Mul, "/" => BinOp::Div,
        "**" => BinOp::Pow, "//" => BinOp::Concat,
        "==" | ".eq." => BinOp::Eq, "/=" | ".ne." => BinOp::NotEq,
        "<" | ".lt." => BinOp::Lt, ">" | ".gt." => BinOp::Gt,
        "<=" | ".le." => BinOp::LtEq, ">=" | ".ge." => BinOp::GtEq,
        ".and." => BinOp::And, ".or." => BinOp::Or,
        ".eqv." => BinOp::Eq, ".neqv." => BinOp::NotEq,
        _ => BinOp::Add,
    }
}

fn meaningful(pair: &Pair<Rule>) -> bool {
    !matches!(pair.as_rule(), Rule::NEWLINE | Rule::EOI)
}

fn is_expr_rule(r: Rule) -> bool {
    matches!(r, Rule::expression | Rule::logical_or | Rule::logical_and | Rule::comparison
        | Rule::addition | Rule::multiplication | Rule::power | Rule::concat | Rule::unary
        | Rule::primary_expr | Rule::literal | Rule::number_literal | Rule::string_literal
        | Rule::identifier | Rule::function_call_or_subscript | Rule::logical_literal | Rule::logical_not)
}

fn to_class_member(stmt: Statement) -> ClassMember {
    match stmt.kind {
        StmtKind::FunctionDecl { .. } => ClassMember::Method(Box::new(stmt)),
        StmtKind::VarDecl { ref declarations, .. } => {
            if let Some(d) = declarations.first() {
                if let BindingPattern::Ident(name) = &d.pattern {
                    return ClassMember::Field {
                        name: name.clone(), type_hint: d.type_hint.clone(), init: d.init.clone(),
                        modifiers: Modifiers::default(), with_events: false, array_bounds: None,
                    };
                }
            }
            ClassMember::Method(Box::new(stmt))
        }
        _ => ClassMember::Method(Box::new(stmt)),
    }
}
