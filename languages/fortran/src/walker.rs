//! Fortran walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.

use super::{FortranParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use vybe_ast::*;
use vybe_compiler::primitives::complex;

const FORTRAN_TBP_IMPL_HANDLE_PREFIX: &str = "__fortran_tbp_impl:";
const FORTRAN_IO_BUFFER_GLOBAL: &str = "__vybe_fortran_io_buffer";

fn to_span(pair: &Pair<Rule>) -> Span {
    let start = pair.as_span().start_pos().line_col();
    let end = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32,
        start_col: start.1 as u32,
        end_line: end.0 as u32,
        end_col: end.1 as u32,
    }
}
const FORTRAN_ARRAY_RESULT_PARAM: &str = "__fortran_array_result";
const FORTRAN_ARRAY_INDEXING: ArrayIndexSemantics = ArrayIndexSemantics::ONE_BASED;

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
                // Copy a whole UTF-8 CHARACTER. `bytes[i] as char` is a Latin-1
                // decode: each byte of `é` became a separate char, so every
                // non-ASCII Fortran source reached the parser as mojibake.
                // unifiedstringplan.md step 0.
                let ch = src[i..].chars().next().expect("index is a char boundary");
                out.push(ch);
                i += ch.len_utf8();
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
                // Whole CHARACTER — see the comment-skip above.
                let ch = src[i..].chars().next().expect("index is a char boundary");
                out.push(ch);
                i += ch.len_utf8();
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
                while j < bytes.len() && bytes[j].is_ascii_digit() {
                    j += 1;
                }
                if j < bytes.len() && (bytes[j] == b'H' || bytes[j] == b'h') {
                    if let Ok(count) = std::str::from_utf8(&bytes[i..j]).unwrap().parse::<usize>() {
                        let after_h = j + 1;
                        let end = (after_h + count).min(bytes.len());
                        if end - after_h == count {
                            let text = &src[after_h..end];
                            // Emit as quoted string with escaped quotes.
                            out.push('"');
                            for ch in text.chars() {
                                if ch == '"' {
                                    out.push('\\');
                                }
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
        // Whole CHARACTER — see the comment-skip above.
        let ch = src[i..].chars().next().expect("index is a char boundary");
        out.push(ch);
        i += ch.len_utf8();
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
            _ => {
                walk_top(pair, &mut name, &mut body, &mut imports)?;
            }
        }
    }

    bind_top_level_type_bound_procedures(&mut body);
    lower_fortran_body_intrinsics(&[], &mut body);
    lower_fortran_array_bounds(&[], &mut body);
    lower_fortran_array_semantics(&[], &mut body);

    repair_remaining_fortran_array_calls(&mut body);
    lower_fortran_array_return_calls(&mut body);
    lower_fortran_array_assignments(&mut body);
    lower_fortran_array_call_arguments(&mut body);
    lower_fortran_array_expressions(&mut body);
    lower_fortran_scalar_array_assignments(&mut body);
    lower_fortran_complex_expressions(&mut body);
    lower_fortran_array_expressions(&mut body);
    lower_fortran_body_intrinsics(&[], &mut body);
    // Last: the dispatch loop rewrites declarations into hoisted ones, and the
    // passes above read declaration shape.
    strip_fortran_value_dummy_markers(&mut body);
    lower_fortran_select_type_specificity(&mut body);
    lower_fortran_external_declarations(&mut body);
    lower_fortran_operator_slots(&mut body);
    lower_fortran_defined_assignment(&mut body);
    lower_fortran_labeled_do(&mut body);
    lower_fortran_integer_division(&mut body);
    lower_fortran_internal_reads(&mut body);
    lower_fortran_c_binding_handles(&mut body);
    lower_fortran_goto_dispatch(&mut body)?;
    body.insert(
        0,
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(FORTRAN_IO_BUFFER_GLOBAL)],
            value: Expression::string(""),
            by_ref: false,
        }),
    );

    Ok(Module {
        canon: Default::default(),
        name,
        language: Lang::Fortran,
        body,
        imports,
        directives: vybe_ast::Directives {
            // `ISHFT`/`SHIFTL`/`SHIFTR` yield ZERO once the count reaches
            // BIT_SIZE — every bit has been shifted out. wasm instead MASKS the
            // count, so the bare instruction answers `ishft(1, 32)` with 1
            // where gfortran answers 0. Nothing in the operands says which rule
            // applies; the language does.
            shift_overflow: Some(vybe_ast::ShiftOverflow::Zero),
            // Fortran identifiers are case-insensitive, ASCII — `PRINT`,
            // `Print` and `print` are one keyword and `myVar`/`MYVAR` one
            // variable. gfortran is the ground truth.
            variable_case: Some(vybe_ast::CaseMatch::Folded),
            callable_case: Some(vybe_ast::CaseMatch::Folded),
            case_alphabet: Some(vybe_ast::CaseAlphabet::Ascii),
            ..Default::default()
        },
    })
}

fn walk_top(
    pair: Pair<Rule>,
    name: &mut String,
    body: &mut Vec<Statement>,
    imports: &mut Vec<Import>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::program_unit => {
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => {
                        if name.is_empty() {
                            *name = p.as_str().to_string();
                        }
                    }
                    _ => {
                        if let Some(st) = walk_stmt(p)? {
                            body.push(st);
                        }
                    }
                }
            }
            lower_fortran_namelist_io(body);
        }
        Rule::module_unit => {
            let mut mname = String::new();
            let mut module_body = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => {
                        if mname.is_empty() {
                            mname = p.as_str().to_string();
                        }
                    }
                    Rule::statement_line => {
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? {
                                module_body.push(st);
                            }
                        }
                    }
                    _ => {
                        if let Some(st) = walk_stmt(p)? {
                            module_body.push(st);
                        }
                    }
                }
            }
            lower_fortran_namelist_io(&mut module_body);
            let module_const_exports = collect_fortran_module_const_exports(&mut module_body);
            body.extend(module_const_exports);
            let members = module_body.into_iter().map(to_class_member).collect();
            body.push(Statement::new(StmtKind::ModuleDecl {
                name: mname,
                members,
                visibility: Visibility::Public,
            }));
        }
        // A submodule holds the BODIES of procedures its parent module only
        // declared, in an `interface` block, as `module function f(…)`. The name
        // belongs to the parent — a submodule is not a scope anybody imports —
        // so the definitions are emitted where that declaration can find them,
        // and the `submodule (parent) name` header itself carries nothing else.
        Rule::submodule_unit => {
            let mut seen_body = false;
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier if !seen_body => {}
                    Rule::statement_line => {
                        seen_body = true;
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? {
                                body.push(st);
                            }
                        }
                    }
                    _ => {
                        seen_body = true;
                        if let Some(st) = walk_stmt(p)? {
                            body.push(st);
                        }
                    }
                }
            }
        }
        Rule::use_statement => {
            let mut parts = pair.into_inner().filter(|p| meaningful(p));
            let mname = parts
                .next()
                .ok_or("missing module name in use")?
                .as_str()
                .to_string();
            let mut names = Vec::new();
            for p in parts {
                if p.as_rule() == Rule::use_name_list {
                    for np in p.into_inner() {
                        if np.as_rule() == Rule::use_name {
                            let mut ni = np.into_inner().filter(|p| meaningful(p));
                            let n = ni
                                .next()
                                .map(|p| p.as_str().to_string())
                                .unwrap_or_default();
                            let a = ni.next().map(|p| p.as_str().to_string());
                            names.push(ImportName { name: n, alias: a });
                        }
                    }
                }
            }
            if names.is_empty() {
                imports.push(Import {
                    kind: ImportKind::Simple {
                        path: mname,
                        alias: None,
                    },
                    span: Span::default(),
                });
            } else {
                imports.push(Import {
                    kind: ImportKind::Named {
                        path: mname,
                        names,
                        level: 0,
                    },
                    span: Span::default(),
                });
            }
        }
        _ => {
            if let Some(st) = walk_stmt(pair)? {
                body.push(st);
            }
        }
    }
    Ok(())
}

fn walk_stmt(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let result = walk_stmt_inner(pair)?;
    Ok(result.map(|mut s| {
        if s.span.start_line == 0 {
            s.span = span;
        }
        s
    }))
}

fn walk_stmt_inner(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::var_declaration => walk_var_decl(pair).map(Some),
        Rule::procedure_decl => walk_procedure_decl(pair).map(Some),
        Rule::assignment_statement => walk_assign(pair).map(Some),
        Rule::where_statement => walk_where(pair).map(Some),
        Rule::call_statement => walk_call(pair).map(Some),
        Rule::if_statement => walk_if(pair).map(Some),
        Rule::do_statement => walk_do(pair).map(Some),
        Rule::do_concurrent_statement => walk_do_concurrent(pair).map(Some),
        Rule::do_while_statement => walk_do_while(pair).map(Some),
        Rule::block_statement => walk_block_construct(pair).map(Some),
        Rule::data_statement => walk_data_statement(pair).map(Some),
        Rule::forall_statement => walk_forall(pair).map(Some),
        // `external f` says f is a PROCEDURE. A companion `integer f` gives its
        // RESULT type, not a variable — but the declaration was emitted as one,
        // so `f()` called the number 0. The names travel as a marker the pass
        // below consumes.
        Rule::external_statement => Ok(Some(Statement::new(StmtKind::Expr(
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(FORTRAN_EXTERNAL_MARKER)),
                args: pair
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::identifier)
                    .map(|p| Argument::positional(Expression::string(p.as_str())))
                    .collect(),
                optional: false,
            }),
        )))),
        Rule::labeled_do_statement => walk_labeled_do(pair).map(Some),
        Rule::select_case_statement => walk_select(pair).map(Some),
        Rule::select_type_statement => walk_select_type(pair).map(Some),
        Rule::select_rank_statement => walk_select_rank(pair).map(Some),
        Rule::enum_statement => walk_enum_statement(pair).map(Some),
        Rule::print_statement => walk_print(pair).map(Some),
        Rule::write_statement => walk_write(pair).map(Some),
        Rule::read_statement => walk_read(pair).map(Some),
        Rule::namelist_statement => walk_namelist_statement(pair).map(Some),
        Rule::subroutine_decl => walk_sub(pair).map(Some),
        Rule::function_decl => walk_func(pair).map(Some),
        Rule::type_decl => walk_type(pair).map(Some),
        Rule::interface_decl | Rule::abstract_interface_decl => walk_interface_decl(pair),
        Rule::allocate_statement => walk_allocate_stmt(pair).map(Some),
        Rule::deallocate_statement => walk_deallocate_stmt(pair).map(Some),
        // `return 1` — the alternate-return selector. `kw_return` is `@{…}`,
        // atomic but NOT silent, so it is a child pair like any other: taking
        // the FIRST meaningful child walked the keyword as the return value and
        // the `1` was never read at all.
        Rule::return_statement => {
            let e = pair
                .into_inner()
                .find(|p| is_expr_rule(p.as_rule()))
                .map(walk_expr)
                .transpose()?;
            Ok(Some(Statement::new(StmtKind::Return(e))))
        }
        Rule::yield_statement => {
            let value = pair
                .into_inner()
                .filter(|p| meaningful(p))
                .next()
                .map(walk_expr)
                .transpose()?;
            Ok(Some(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Yield(value.map(Box::new)),
            )))))
        }
        // `cycle` / `exit` take an optional CONSTRUCT NAME, and the name is the
        // whole point of naming a loop — `cycle outer` from an inner loop must
        // reach the outer one. The name was being dropped, so both always
        // lowered to `Implicit` and hit the innermost loop.
        Rule::cycle_statement => Ok(Some(Statement::new(StmtKind::Continue(
            fortran_loop_target_name(&pair).map_or(ContinueTarget::Implicit, ContinueTarget::Label),
        )))),
        Rule::exit_statement => Ok(Some(Statement::new(StmtKind::Break(
            fortran_loop_target_name(&pair).map_or(BreakTarget::Implicit, BreakTarget::Label),
        )))),
        // `STOP <code>` sets the process status; a bare `STOP` is ordinary
        // termination. The code was being dropped here — every `stop 1` exited
        // 0, so a Fortran program could not report failure at all (gfortran
        // gives 1). The grammar already captures it.
        Rule::stop_statement | Rule::error_stop_statement => {
            let code = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::expression)
                .map(walk_expr)
                .transpose()?;
            match code {
                Some(code) => Ok(Some(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident("__fortran_stop")),
                        args: vec![Argument::positional(code)],
                        optional: false,
                    },
                ))))),
                None => Ok(Some(Statement::new(StmtKind::Return(None)))),
            }
        }
        Rule::expression_statement => {
            let e = walk_expr(pair.into_inner().next().ok_or("empty expr")?)?;
            if let Some(stmt) = lower_intrinsic_statement(&e) {
                return Ok(Some(stmt));
            }
            Ok(Some(Statement::new(StmtKind::Expr(e))))
        }
        Rule::goto_statement => Ok(Some(fortran_goto_marker_statement(
            fortran_first_statement_label(&pair).ok_or("missing goto label")?,
        ))),
        Rule::computed_goto_statement => walk_computed_goto(pair).map(Some),
        Rule::arithmetic_if_statement => walk_arithmetic_if(pair).map(Some),
        // `CONTINUE` is a no-op, but a LABELLED one is a goto target, and the
        // label lives on the enclosing `statement_line`. An empty block keeps
        // the line alive so the label survives to the dispatch pass.
        Rule::continue_statement => Ok(Some(Statement::new(StmtKind::Block(Vec::new())))),
        Rule::statement_line => {
            let mut stmts = walk_statement_line_stmts(pair)?;
            match stmts.len() {
                0 => Ok(None),
                1 => Ok(Some(stmts.remove(0))),
                _ => Ok(Some(Statement::new(StmtKind::Block(stmts)))),
            }
        }
        Rule::use_statement => Ok(None),
        Rule::program_unit | Rule::module_unit => {
            let mut body = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if p.as_rule() == Rule::statement_line {
                    for s in p.into_inner().filter(|p| meaningful(p)) {
                        if let Some(st) = walk_stmt(s)? {
                            body.push(st);
                        }
                    }
                } else if p.as_rule() != Rule::identifier {
                    if let Some(st) = walk_stmt(p)? {
                        body.push(st);
                    }
                }
            }
            lower_fortran_namelist_io(&mut body);
            Ok(Some(Statement::new(StmtKind::Block(body))))
        }
        _ => Ok(None),
    }
}

fn walk_body<'a>(pairs: impl Iterator<Item = Pair<'a, Rule>>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pairs {
        match p.as_rule() {
            // Whole-line, not per-statement: a procedure body is a place where
            // GOTO targets live, and the label sits on the LINE.
            Rule::statement_line => body.extend(walk_statement_line_stmts(p)?),
            Rule::identifier => {}
            _ => {
                if let Some(st) = walk_stmt(p)? {
                    body.push(st);
                }
            }
        }
    }
    Ok(body)
}

fn walk_inline_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt in pair.into_inner().filter(|p| meaningful(p)) {
        if let Some(lowered) = walk_stmt(stmt)? {
            body.push(lowered);
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

fn parse_derived_type_name(type_hint: &str) -> Option<String> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let prefix_len = if lower.starts_with("type(") {
        5
    } else if lower.starts_with("class(") {
        6
    } else {
        return None;
    };

    let suffix = &trimmed[prefix_len..];
    let end = suffix.find(')')?;
    let name = suffix[..end].trim();
    if name.is_empty() || name == "*" {
        return None;
    }
    Some(name.to_string())
}

/// The declared type hint, carrying the `pointer` attribute when there is a
/// derived type to point AT.
///
/// A `pointer` component refers to a derived type without storing one, and the
/// shared spelling for that is a `*` prefix — `type_hint_stores_by_value`
/// declines `*T`, `^T`, `[]T` alike. Dropping the attribute made
/// `type(node), pointer :: next` indistinguishable from a stored `type(node)`,
/// so the class ctor default-constructed the component
/// (`classes.rs` → `user_value_type_name_from_hint` → `__node_ctor_0`) and a
/// SELF-referential node type recursed until the stack died.
///
/// Only derived types are marked. `integer, pointer :: p` names no class to
/// construct, so prefixing it would change a spelling nothing reads.
///
/// This is the same fact `walk_var_decl` already uses to suppress the implicit
/// `New` for a pointer variable — a pointer starts unassociated. It was simply
/// never written down anywhere the class-member path could read it.
fn fortran_pointer_type_hint(type_hint: Option<&str>, is_pointer: bool) -> Option<String> {
    let hint = type_hint?;
    if !is_pointer || parse_derived_type_name(hint).is_none() {
        return Some(hint.to_string());
    }
    Some(format!("*{}", hint.trim()))
}

/// `type(c_ptr)` / `type(c_funptr)` — the `iso_c_binding` opaque handles.
fn is_fortran_opaque_c_handle(type_hint: &str) -> bool {
    // Either spelling reaches here: the written `type(c_ptr)` and the bare
    // `c_ptr` a earlier normalization may already have reduced it to.
    let name = parse_derived_type_name(type_hint)
        .unwrap_or_else(|| type_hint.trim().trim_end_matches("()").trim().to_string());
    matches!(
        name.to_ascii_lowercase().as_str(),
        "c_ptr" | "c_funptr" | "c_devptr"
    )
}

fn fortran_type_hint_array_rank(type_hint: &str) -> usize {
    let mut rank = 0;
    let mut rest = type_hint.trim_end();
    while let Some(stripped) = rest.strip_suffix("()") {
        rank += 1;
        rest = stripped.trim_end();
    }
    rank
}

fn strip_fortran_type_hint_array_rank(type_hint: &str) -> &str {
    let mut rest = type_hint.trim_end();
    while let Some(stripped) = rest.strip_suffix("()") {
        rest = stripped.trim_end();
    }
    rest
}

fn fortran_declared_array_rank(array_bounds: Option<&[Expression]>) -> usize {
    match array_bounds {
        Some(bounds) if bounds.is_empty() => 1,
        Some(bounds) => bounds.len(),
        None => 0,
    }
}

fn fortran_array_type_hint(type_hint: &str, array_bounds: Option<&[Expression]>) -> String {
    let rank =
        fortran_declared_array_rank(array_bounds).max(fortran_type_hint_array_rank(type_hint));
    if rank == 0 {
        return strip_fortran_type_hint_array_rank(type_hint).to_string();
    }
    format!(
        "{}{}",
        strip_fortran_type_hint_array_rank(type_hint),
        "()".repeat(rank)
    )
}

/// Per-dimension declaration facts: the extent expression, plus the declared
/// lower bound when the spec spelled one (`v(-2:3)`). A dimension written as a
/// bare extent (`v(5)`) has Fortran's default origin of 1 and reports `None`.
struct FortranDimensionSpecs {
    extents: Vec<Expression>,
    lower_bounds: Vec<Option<Expression>>,
}

fn parse_fortran_dimension_spec_list(
    pair: Pair<Rule>,
) -> Result<FortranDimensionSpecs, String> {
    let mut dim_bounds = Vec::new();
    let mut lower_bounds = Vec::new();
    for spec in pair.into_inner().filter(|p| meaningful(p)) {
        if spec.as_rule() != Rule::dimension_spec {
            continue;
        }
        let exprs: Vec<Pair<Rule>> = spec
            .clone()
            .into_inner()
            .filter(|p| meaningful(p))
            .collect();
        let this_size = if exprs.len() == 1 {
            lower_bounds.push(None);
            walk_expr(exprs.into_iter().next().unwrap())?
        } else if exprs.len() == 2 {
            let lo = walk_expr(exprs[0].clone())?;
            let hi = walk_expr(exprs[1].clone())?;
            lower_bounds.push(Some(lo.clone()));
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
            continue;
        };
        dim_bounds.push(this_size);
    }
    Ok(FortranDimensionSpecs {
        extents: dim_bounds,
        lower_bounds,
    })
}

/// Companion vector holding an array's declared per-dimension lower bounds.
/// Fortran's array descriptor carries the origin alongside the extents; the AST
/// declaration records only extents, so the origin is declared beside the array.
fn fortran_origin_variable_name(array: &str) -> String {
    format!("__fortran_origin_{}", array.to_ascii_lowercase())
}

/// `None` when every dimension uses Fortran's default origin of 1 — the usual
/// case, which needs no descriptor because 1 is what everything already assumes.
fn fortran_origin_declarator(
    name: &str,
    lower_bounds: &[Option<Expression>],
) -> Option<VarDeclarator> {
    if !lower_bounds.iter().any(|bound| {
        bound
            .as_ref()
            .is_some_and(|bound| !matches!(bound.kind, ExprKind::Lit(Literal::Int(1))))
    }) {
        return None;
    }
    let origins = lower_bounds
        .iter()
        .map(|bound| ArrayElement {
            key: None,
            value: bound.clone().unwrap_or_else(|| Expression::int(1)),
            spread: false,
            by_ref: false,
        })
        .collect();
    Some(VarDeclarator {
        pattern: BindingPattern::Ident(fortran_origin_variable_name(name)),
        type_hint: None,
        init: Some(Expression::new(ExprKind::Array(origins))),
        array_bounds: None,
        with_events: false,
    })
}

/// What a declaration told us about an array's shape. `origins` is empty when
/// every dimension uses Fortran's default origin of 1; `extents` is empty for a
/// deferred-shape array whose size only exists at run time.
#[derive(Default, Clone)]
struct FortranArrayShape {
    origins: Vec<Expression>,
    extents: Vec<Expression>,
}

#[derive(Default, Clone)]
struct FortranBoundsEnv {
    shapes: HashMap<String, FortranArrayShape>,
}

impl FortranBoundsEnv {
    fn shape_of(&self, expr: &Expression) -> Option<&FortranArrayShape> {
        let ExprKind::Ident(name) = &expr.kind else {
            // An array SECTION, a function result, or an assumed-shape dummy has
            // origin 1 by definition — the declared origin never travels with the
            // value, only with the name it was declared under.
            return None;
        };
        self.shapes.get(&name.to_ascii_lowercase())
    }

    fn rank_of(&self, expr: &Expression) -> Option<usize> {
        let shape = self.shape_of(expr)?;
        let rank = shape.extents.len().max(shape.origins.len());
        (rank > 0).then_some(rank)
    }

    /// Lower bound of `expr` along 1-based `dim`.
    fn origin(&self, expr: &Expression, dim: usize) -> Expression {
        self.shape_of(expr)
            .and_then(|shape| shape.origins.get(dim - 1))
            .cloned()
            .unwrap_or_else(|| Expression::int(1))
    }

    /// Extent of `expr` along 1-based `dim`, from the declaration when it stated
    /// one and from the value itself otherwise. `None` when neither can answer:
    /// reading dimension 2 off a value means indexing into it, and an array of
    /// unknown rank — an assumed-rank `a(..)`, a `mold=` allocation — may have no
    /// second dimension to index.
    fn extent(&self, expr: &Expression, dim: usize) -> Option<Expression> {
        if let Some(extent) = self
            .shape_of(expr)
            .and_then(|shape| shape.extents.get(dim - 1))
        {
            return Some(extent.clone());
        }
        if dim > 1 && self.rank_of(expr).is_none_or(|rank| rank < dim) {
            return None;
        }
        // Rank-N arrays are nested, so dimension `d` is the length of the array
        // reached by indexing the first element `d - 1` times.
        let mut target = expr.clone();
        for _ in 1..dim {
            target = Expression::new(ExprKind::Index {
                object: Box::new(target),
                index: Box::new(Expression::int(0)),
                null_safe: false,
            });
        }
        Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("size")),
            args: vec![Argument::positional(target)],
            optional: false,
        }))
    }

    fn record(&mut self, name: &str, shape: FortranArrayShape) {
        let key = name.to_ascii_lowercase();
        let entry = self.shapes.entry(key).or_default();
        if !shape.origins.is_empty() {
            entry.origins = shape.origins;
        }
        if !shape.extents.is_empty() {
            entry.extents = shape.extents;
        }
    }
}

fn fortran_literal_dim(args: &[Argument]) -> Option<usize> {
    let dim = args.iter().find(|arg| {
        arg.name
            .as_deref()
            .is_none_or(|name| name.eq_ignore_ascii_case("dim"))
    })?;
    match dim.value.kind {
        ExprKind::Lit(Literal::Int(value)) if value >= 1 => Some(value as usize),
        _ => None,
    }
}

/// `lbound`/`ubound`/`size(a, dim)` — the intrinsics that read an array's
/// declared bounds. They answer from the declaration, so they cannot be folded
/// where the other intrinsics are: nothing at that point knows the shape.
fn lower_fortran_array_bounds(params: &[Param], body: &mut [Statement]) {
    let mut env = FortranBoundsEnv::default();
    for param in params {
        // A dummy argument's bounds are its own: an assumed-shape array starts at
        // 1 whatever the caller declared, so the rank is all that carries over.
        let rank = param
            .type_hint
            .as_deref()
            .map(fortran_type_hint_array_rank)
            .unwrap_or(0);
        if rank > 0 {
            env.record(
                &param.name,
                FortranArrayShape {
                    origins: vec![Expression::int(1); rank],
                    extents: Vec::new(),
                },
            );
        }
    }
    lower_fortran_array_bounds_with_env(body, &mut env);
}

fn lower_fortran_array_bounds_with_env(body: &mut [Statement], env: &mut FortranBoundsEnv) {
    for statement in body.iter_mut() {
        record_fortran_array_shapes(statement, env);
        rewrite_fortran_bounds_in_statement(statement, env);
        match &mut statement.kind {
            // Contained procedures are lowered by their own walker, before this
            // body is assembled — visiting them again would shift every subscript
            // a second time.
            StmtKind::FunctionDecl { .. } => {}
            // A `Block` here is a statement GROUP the walker synthesised — an
            // `allocate` with its descriptor update, say — not a Fortran scope,
            // so what it declares stays visible to the statements after it.
            // A named construct wraps its statement in `Labeled`; the wrapper is
            // transparent, so the env passes through to whatever it names.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_bounds_with_env(std::slice::from_mut(body.as_mut()), env)
            }
            StmtKind::Block(stmts) => lower_fortran_array_bounds_with_env(stmts, env),
            StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested = env.clone();
                lower_fortran_array_bounds_with_env(stmts, &mut nested);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut nested = env.clone();
                lower_fortran_array_bounds_with_env(then_body, &mut nested);
                for (cond, elif_body) in elifs {
                    rewrite_fortran_bounds_in_expr(cond, env);
                    let mut nested = env.clone();
                    lower_fortran_array_bounds_with_env(elif_body, &mut nested);
                }
                if let Some(else_body) = else_body {
                    let mut nested = env.clone();
                    lower_fortran_array_bounds_with_env(else_body, &mut nested);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut nested = env.clone();
                lower_fortran_array_bounds_with_env(stmts, &mut nested);
                if let Some(else_body) = else_body {
                    let mut nested = env.clone();
                    lower_fortran_array_bounds_with_env(else_body, &mut nested);
                }
            }
            StmtKind::For { body: stmts, .. } => {
                let mut nested = env.clone();
                lower_fortran_array_bounds_with_env(stmts, &mut nested);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases.iter_mut() {
                    let mut nested = env.clone();
                    lower_fortran_array_bounds_with_env(&mut case.body, &mut nested);
                }
                if let Some(default) = default {
                    let mut nested = env.clone();
                    lower_fortran_array_bounds_with_env(default, &mut nested);
                }
            }
            _ => {}
        }
    }
}

fn record_fortran_array_shapes(statement: &Statement, env: &mut FortranBoundsEnv) {
    match &statement.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if let Some(array) = name
                    .to_ascii_lowercase()
                    .strip_prefix("__fortran_origin_")
                    .map(str::to_string)
                {
                    if let Some(origins) = declaration.init.as_ref().and_then(fortran_array_literal)
                    {
                        env.record(
                            &array,
                            FortranArrayShape {
                                origins,
                                extents: Vec::new(),
                            },
                        );
                    }
                    continue;
                }
                if let Some(extents) = declaration
                    .array_bounds
                    .as_ref()
                    .filter(|bounds| !bounds.is_empty())
                {
                    env.record(
                        name,
                        FortranArrayShape {
                            origins: Vec::new(),
                            extents: extents.clone(),
                        },
                    );
                }
            }
        }
        StmtKind::Expr(Expression {
            kind: ExprKind::Call { callee, args, .. },
            ..
        }) => {
            if !matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("allocate"))
            {
                return;
            }
            for arg in args {
                let ExprKind::Call {
                    callee: target,
                    args: dims,
                    ..
                } = &arg.value.kind
                else {
                    continue;
                };
                let ExprKind::Ident(name) = &target.kind else {
                    continue;
                };
                if dims.is_empty() {
                    continue;
                }
                env.record(
                    name,
                    FortranArrayShape {
                        origins: Vec::new(),
                        extents: dims.iter().map(|dim| dim.value.clone()).collect(),
                    },
                );
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            let [target] = targets.as_slice() else {
                return;
            };
            let ExprKind::Ident(name) = &target.kind else {
                return;
            };
            let Some(array) = name
                .to_ascii_lowercase()
                .strip_prefix("__fortran_origin_")
                .map(str::to_string)
            else {
                return;
            };
            if let Some(origins) = fortran_array_literal(value) {
                env.record(
                    &array,
                    FortranArrayShape {
                        origins,
                        extents: Vec::new(),
                    },
                );
            }
        }
        _ => {}
    }
}

fn rewrite_fortran_bounds_in_statement(statement: &mut Statement, env: &FortranBoundsEnv) {
    match &mut statement.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        }
        | StmtKind::If { cond: expr, .. }
        | StmtKind::While { cond: expr, .. }
        | StmtKind::DoWhile { cond: expr, .. }
        | StmtKind::ForIn { iter: expr, .. }
        | StmtKind::Switch { expr, .. } => rewrite_fortran_bounds_in_expr(expr, env),
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_fortran_bounds_in_expr(target, env);
            }
            rewrite_fortran_bounds_in_expr(value, env);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_fortran_bounds_in_expr(target, env);
            rewrite_fortran_bounds_in_expr(value, env);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations {
                if let Some(init) = &mut declaration.init {
                    rewrite_fortran_bounds_in_expr(init, env);
                }
            }
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_fortran_bounds_in_expr(cond, env);
            }
            if let Some(update) = update {
                rewrite_fortran_bounds_in_expr(update, env);
            }
        }
        _ => {}
    }
}

fn rewrite_fortran_bounds_in_expr(expr: &mut Expression, env: &FortranBoundsEnv) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_fortran_bounds_in_expr(left, env);
            rewrite_fortran_bounds_in_expr(right, env);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Member { object: inner, .. } => rewrite_fortran_bounds_in_expr(inner, env),
        ExprKind::ArrayMap { array, body, .. } => {
            rewrite_fortran_bounds_in_expr(array, env);
            rewrite_fortran_bounds_in_expr(body, env);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_fortran_bounds_in_expr(cond, env);
            rewrite_fortran_bounds_in_expr(then, env);
            rewrite_fortran_bounds_in_expr(else_, env);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_fortran_bounds_in_expr(object, env);
            rewrite_fortran_bounds_in_expr(index, env);
        }
        ExprKind::Slice { lower, upper, step } => {
            for part in [lower, upper, step].into_iter().flatten() {
                rewrite_fortran_bounds_in_expr(part, env);
            }
        }
        ExprKind::Call { callee, args, .. } | ExprKind::New { class: callee, args } => {
            // `allocate(v(-4:1))` spells bounds, not subscripts — its arguments
            // describe the array being made and must survive untouched.
            if matches!(&callee.kind, ExprKind::Ident(name)
                if name.eq_ignore_ascii_case("allocate") || name.eq_ignore_ascii_case("deallocate"))
            {
                return;
            }
            rewrite_fortran_bounds_in_expr(callee, env);
            for arg in args {
                rewrite_fortran_bounds_in_expr(&mut arg.value, env);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_fortran_bounds_in_expr(target, env);
            rewrite_fortran_bounds_in_expr(value, env);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_fortran_bounds_in_expr(&mut item.value, env);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_fortran_bounds_in_expr(item, env);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(inner) = part {
                    rewrite_fortran_bounds_in_expr(inner, env);
                }
            }
        }
        _ => {}
    }

    if let Some(folded) = fold_fortran_bounds_call(expr, env) {
        *expr = folded;
        return;
    }
    offset_fortran_subscripts(expr, env);
}

/// `integer :: v(-2:3)` numbers its elements from −2, so `v(i)` is not the i-th
/// slot. Shifting the subscript back to a 1-based one here — while it is still a
/// Fortran subscript — leaves the generic 0-based lowering downstream untouched.
fn offset_fortran_subscripts(expr: &mut Expression, env: &FortranBoundsEnv) {
    match &mut expr.kind {
        // A read — `v(i)` still spelled as a call, before the generic subscript
        // lowering turns it into an index chain.
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Ident(name) = &callee.kind else {
                return;
            };
            let Some(origins) = fortran_declared_origins(env, name) else {
                return;
            };
            for (dim, arg) in args.iter_mut().enumerate() {
                let Some(origin) = origins.get(dim) else {
                    continue;
                };
                shift_fortran_subscript_operand(&mut arg.value, origin);
            }
        }
        // An assignment target — the walker builds those as an index chain
        // directly, one node per subscript.
        ExprKind::Index { object, index, .. } => {
            let Some((name, depth)) = fortran_index_chain_base(object) else {
                return;
            };
            let Some(origins) = fortran_declared_origins(env, &name) else {
                return;
            };
            let Some(origin) = origins.get(depth) else {
                return;
            };
            shift_fortran_subscript_operand(index, origin);
        }
        _ => {}
    }
}

fn fortran_declared_origins(env: &FortranBoundsEnv, name: &str) -> Option<Vec<Expression>> {
    let origins = &env.shapes.get(&name.to_ascii_lowercase())?.origins;
    (!origins.is_empty()).then(|| origins.clone())
}

/// The name an index chain is rooted at, and how many subscripts have already
/// been applied to it — which is the 0-based dimension of the next one.
fn fortran_index_chain_base(expr: &Expression) -> Option<(String, usize)> {
    match &expr.kind {
        ExprKind::Ident(name) => Some((name.clone(), 0)),
        ExprKind::Index { object, .. } => {
            let (name, depth) = fortran_index_chain_base(object)?;
            Some((name, depth + 1))
        }
        _ => None,
    }
}

fn shift_fortran_subscript_operand(operand: &mut Expression, origin: &Expression) {
    if matches!(origin.kind, ExprKind::Lit(Literal::Int(1))) {
        return;
    }
    match &mut operand.kind {
        ExprKind::Slice { lower, upper, .. } => {
            for part in [lower, upper].into_iter().flatten() {
                **part = fortran_shift_subscript(part, origin);
            }
        }
        _ => *operand = fortran_shift_subscript(operand, origin),
    }
}

fn fortran_shift_subscript(subscript: &Expression, origin: &Expression) -> Expression {
    let shift = fortran_add_ints(origin.clone(), Expression::int(-1));
    if matches!(shift.kind, ExprKind::Lit(Literal::Int(0))) {
        return subscript.clone();
    }
    if let (ExprKind::Lit(Literal::Int(index)), ExprKind::Lit(Literal::Int(shift))) =
        (&subscript.kind, &shift.kind)
    {
        return Expression::int(index - shift);
    }
    Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(subscript.clone()),
        right: Box::new(shift),
    })
}

fn fold_fortran_bounds_call(expr: &Expression, env: &FortranBoundsEnv) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let array = &args.first()?.value;
    let name = name.to_ascii_lowercase();
    let upper = match name.as_str() {
        "lbound" => false,
        "ubound" => true,
        // `size(a)` already answers through the array-length builtin; only the
        // per-dimension form needs the declaration.
        "size" if args.len() > 1 => {
            let dim = fortran_literal_dim(&args[1..])?;
            return env.extent(array, dim);
        }
        // `shape(a)` is the EXTENT vector — `ubound - lbound + 1` per dimension,
        // which the descriptor already holds directly. It reads the same
        // declaration `lbound`/`ubound` do, so it belongs here rather than as a
        // builtin: the extents are a compile-time fact, and a runtime helper
        // would have to rediscover them by walking the nest.
        //
        // `shape(a, dim)` is not a Fortran form — `size(a, dim)` is the scalar
        // ask — so the vector is the only shape to build.
        "shape" => {
            let rank = env.rank_of(array)?;
            let extents = (1..=rank)
                .map(|dim| {
                    Some(ArrayElement {
                        key: None,
                        value: env.extent(array, dim)?,
                        spread: false,
                        by_ref: false,
                    })
                })
                .collect::<Option<Vec<_>>>()?;
            return Some(Expression::new(ExprKind::Array(extents)));
        }
        _ => return None,
    };
    let bound = |dim: usize| {
        let origin = env.origin(array, dim);
        if !upper {
            return Some(origin);
        }
        Some(fortran_add_ints(
            origin,
            fortran_add_ints(env.extent(array, dim)?, Expression::int(-1)),
        ))
    };
    if args.len() > 1 {
        return bound(fortran_literal_dim(&args[1..])?);
    }
    // No `dim`: the answer is a vector with one entry per dimension, so the rank
    // has to be known. An unranked array leaves the call alone rather than
    // guessing a shape.
    let rank = env.rank_of(array)?;
    let bounds = (1..=rank)
        .map(|dim| {
            Some(ArrayElement {
                key: None,
                value: bound(dim)?,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Option<Vec<_>>>()?;
    Some(Expression::new(ExprKind::Array(bounds)))
}

fn fortran_add_ints(left: Expression, right: Expression) -> Expression {
    if let (ExprKind::Lit(Literal::Int(a)), ExprKind::Lit(Literal::Int(b))) =
        (&left.kind, &right.kind)
    {
        return Expression::int(a + b);
    }
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn fortran_array_literal(expr: &Expression) -> Option<Vec<Expression>> {
    let ExprKind::Array(elements) = &expr.kind else {
        return None;
    };
    if elements.iter().any(|element| element.spread) {
        return None;
    }
    Some(
        elements
            .iter()
            .map(|element| element.value.clone())
            .collect(),
    )
}

fn has_deferred_fortran_dimension_spec(pair: &Pair<Rule>) -> bool {
    pair.clone()
        .into_inner()
        .filter(|p| meaningful(p))
        .any(|spec| {
            spec.as_rule() == Rule::dimension_spec
                && spec.into_inner().filter(|p| meaningful(p)).next().is_none()
        })
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner();
    let type_hint = inner.next().map(|p| p.as_str().trim().to_string());
    let mut is_pointer = false;
    let mut is_value = false;
    let mut is_allocatable = false;
    let mut has_intent = false;
    let mut attr_dim_bounds: Vec<Expression> = Vec::new();
    let mut attr_dim_lower_bounds: Vec<Option<Expression>> = Vec::new();
    let mut has_attr_array_bounds = false;
    let mut has_attr_deferred_array_bounds = false;
    let mut trailing = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::var_attributes => {
                for attr in p.into_inner().filter(|attr| meaningful(attr)) {
                    match attr.as_rule() {
                        Rule::var_attribute => {
                            let attr_text = attr.as_str().trim().to_ascii_lowercase();
                            if attr_text == "value" {
                                is_value = true;
                            } else if attr_text == "pointer" {
                                is_pointer = true;
                            } else if attr_text == "allocatable" {
                                is_allocatable = true;
                            } else if attr_text.starts_with("intent(") {
                                has_intent = true;
                            }
                            for child in attr.into_inner().filter(|child| meaningful(child)) {
                                if child.as_rule() == Rule::dimension_spec_list {
                                    if has_deferred_fortran_dimension_spec(&child) {
                                        has_attr_deferred_array_bounds = true;
                                        has_attr_array_bounds = true;
                                    }
                                    let specs = parse_fortran_dimension_spec_list(child)?;
                                    if !specs.extents.is_empty() {
                                        attr_dim_bounds = specs.extents;
                                        attr_dim_lower_bounds = specs.lower_bounds;
                                        has_attr_array_bounds = true;
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => trailing.push(p),
        }
    }
    let mut declarations = Vec::new();
    for p in trailing {
        if p.as_rule() == Rule::var_declarator_list {
            for d in p.into_inner() {
                if d.as_rule() == Rule::var_declarator {
                    let mut di = d.into_inner().filter(|p| meaningful(p));
                    let nm = di
                        .next()
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    let mut init = None;
                    let mut dim_bounds: Vec<Expression> = attr_dim_bounds.clone();
                    let mut dim_lower_bounds: Vec<Option<Expression>> =
                        attr_dim_lower_bounds.clone();
                    let mut has_array_bounds =
                        has_attr_array_bounds || has_attr_deferred_array_bounds;
                    for pp in di {
                        match pp.as_rule() {
                            Rule::dimension_spec_list => {
                                if has_deferred_fortran_dimension_spec(&pp) {
                                    has_array_bounds = true;
                                    dim_bounds.clear();
                                    dim_lower_bounds.clear();
                                }
                                let specs = parse_fortran_dimension_spec_list(pp)?;
                                if !specs.extents.is_empty() {
                                    has_array_bounds = true;
                                    dim_bounds = specs.extents;
                                    dim_lower_bounds = specs.lower_bounds;
                                }
                            }
                            Rule::codimension_spec_list => { /* PGAS — model as scalar/array */ }
                            _ => {
                                init = Some(walk_expr(pp)?);
                            }
                        }
                    }
                    // Fixed Fortran arrays should be initialized from their declared
                    // per-dimension bounds in the compiler so multidimensional arrays
                    // allocate nested runtime arrays instead of a single flat buffer.
                    if init.is_none() {
                        if !has_array_bounds && !is_pointer && !is_allocatable && !has_intent {
                            if let Some(class_name) =
                                type_hint.as_deref().and_then(parse_derived_type_name)
                            {
                                // `type(c_ptr)` LOOKS like a derived type and is
                                // not one — it is an opaque `iso_c_binding`
                                // handle with no components and no constructor,
                                // and it starts as null. Constructing it emitted
                                // `new c_ptr()` against a class that does not
                                // exist, so the DECLARATION died on "undefined
                                // is not callable" before any C call ran.
                                init = Some(
                                    if type_hint
                                        .as_deref()
                                        .is_some_and(is_fortran_opaque_c_handle)
                                    {
                                        Expression::new(ExprKind::Lit(Literal::Null))
                                    } else {
                                        Expression::new(ExprKind::New {
                                            class: Box::new(Expression::new(ExprKind::Ident(
                                                class_name,
                                            ))),
                                            args: Vec::new(),
                                        })
                                    },
                                );
                            }
                        }
                    }
                    // Fortran `character(len=N) :: s = 'literal'` — pad the literal
                    // with trailing blanks so `len(s)` returns the declared length.
                    // Pure JS-shape rewrite: wrap the init in `s.padEnd(N, ' ')`.
                    if !has_array_bounds {
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
                                            Argument::positional(Expression::new(ExprKind::Lit(
                                                Literal::Int(declared_len),
                                            ))),
                                            Argument::positional(Expression::new(ExprKind::Lit(
                                                Literal::Str(" ".into()),
                                            ))),
                                        ],
                                        optional: false,
                                    });
                                    init = Some(padded);
                                } else {
                                    // No init — synthesize a string of N spaces so len(s) returns N.
                                    let spaces: String = " ".repeat(declared_len as usize);
                                    init =
                                        Some(Expression::new(ExprKind::Lit(Literal::Str(spaces))));
                                }
                            }
                        }
                    }
                    // A declared origin (`v(-2:3)`) is a fact about the array that
                    // nothing downstream can recover: `array_bounds` records the
                    // EXTENT, `hi - lo + 1`, and the origin is gone. Fortran carries
                    // it in the array descriptor, so declare it as one — a companion
                    // vector of per-dimension lower bounds, in the same scope, which
                    // `lbound`/`ubound`/`size(a, dim)` and subscripting read back.
                    let declared_hint =
                        fortran_pointer_type_hint(type_hint.as_deref(), is_pointer);
                    if let Some(origin) = fortran_origin_declarator(&nm, &dim_lower_bounds) {
                        declarations.push(VarDeclarator {
                            pattern: BindingPattern::Ident(nm.clone()),
                            type_hint: declared_hint.clone().map(Into::into),
                            init,
                            array_bounds: has_array_bounds.then_some(dim_bounds),
                            with_events: false,
                        });
                        declarations.push(origin);
                        continue;
                    }
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(nm),
                        type_hint: declared_hint.map(Into::into),
                        init,
                        array_bounds: has_array_bounds.then_some(dim_bounds),
                        with_events: false,
                    });
                }
            }
        }
    }
    // `integer, value :: x` — a VALUE dummy is passed BY VALUE, so assigning to
    // it must not reach the caller. Fortran's default is by reference and the
    // promotion pass turns any mutated dummy into an alias; this marker is how
    // that pass learns which dummies are exempt.
    if is_value {
        let names: Vec<Argument> = declarations
            .iter()
            .filter_map(|declaration| match &declaration.pattern {
                BindingPattern::Ident(name) => Some(Argument::positional(Expression::string(name))),
                _ => None,
            })
            .collect();
        if !names.is_empty() {
            return Ok(Statement::new(StmtKind::Block(vec![
                Statement::new(StmtKind::VarDecl {
                    declarations,
                    kind: VarDeclKind::Dim,
                }),
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(FORTRAN_VALUE_DUMMY_MARKER)),
                    args: names,
                    optional: false,
                }))),
            ])));
        }
    }
    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Dim,
    }))
}

fn walk_enum_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut statements = Vec::new();
    let mut next_value = 0_i64;
    let mut known_values: HashMap<String, i64> = HashMap::new();

    for decl in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::enumerator_decl)
    {
        for value_pair in decl
            .into_inner()
            .filter(|p| p.as_rule() == Rule::enumerator_value)
        {
            let mut inner = value_pair.into_inner().filter(|p| meaningful(p));
            let Some(name_pair) = inner.next() else {
                continue;
            };
            let name = name_pair.as_str().to_string();
            let init = if let Some(expr_pair) = inner.next() {
                walk_expr(expr_pair)?
            } else {
                Expression::new(ExprKind::Lit(Literal::Int(next_value)))
            };
            if let Some(value) = fortran_const_int_expr(&init, &known_values) {
                next_value = value + 1;
                known_values.insert(name.clone(), value);
            } else {
                next_value += 1;
            }
            statements.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: Some("integer".to_string().into()),
                    init: Some(init),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }));
        }
    }

    Ok(Statement::new(StmtKind::Block(statements)))
}

/// The module's data, lifted to top-level declarations.
///
/// A Fortran module variable has STATIC storage and one instance shared by
/// every unit that `use`s it — writing `seen` in a subroutine and reading it in
/// the program must see the same object. Only `parameter` (Const) members were
/// lifted, so a plain `integer :: seen = 0` stayed inside the `ModuleDecl` and
/// `use stash` never brought it into scope: the write and the read landed on
/// different things and the read always returned the declared initial value.
///
/// Lifting is not enough on its own: while the `ModuleDecl` still declared the
/// same name, the namespace object won the READ (`struct.get seen`) and the
/// bare global took the WRITE (`global.set seen`). The declaration is therefore
/// MOVED, not copied — Fortran has no `module%var` spelling, so nothing needs
/// the namespace member.
fn collect_fortran_module_const_exports(statements: &mut Vec<Statement>) -> Vec<Statement> {
    let mut exports = Vec::new();
    for stmt in statements.iter_mut() {
        if let StmtKind::Block(items) = &mut stmt.kind {
            let mut kept = Vec::with_capacity(items.len());
            for item in std::mem::take(items) {
                match item.kind {
                    StmtKind::VarDecl { .. } => exports.push(item),
                    _ => kept.push(item),
                }
            }
            *items = kept;
        }
    }
    let mut kept = Vec::with_capacity(statements.len());
    for stmt in std::mem::take(statements) {
        match stmt.kind {
            StmtKind::VarDecl { .. } => exports.push(stmt),
            _ => kept.push(stmt),
        }
    }
    *statements = kept;
    exports
}

fn fortran_const_int_expr(expr: &Expression, known_values: &HashMap<String, i64>) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) => Some(*value as i64),
        ExprKind::Ident(name) => known_values.get(name).copied(),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => fortran_const_int_expr(expr, known_values).map(|value| -value),
        ExprKind::Unary {
            op: UnaryOp::Pos,
            expr,
        } => fortran_const_int_expr(expr, known_values),
        ExprKind::Binary { op, left, right } => {
            let left = fortran_const_int_expr(left, known_values)?;
            let right = fortran_const_int_expr(right, known_values)?;
            match op {
                BinOp::Add => Some(left + right),
                BinOp::Sub => Some(left - right),
                BinOp::Mul => Some(left * right),
                BinOp::Div if right != 0 => Some(left / right),
                BinOp::Mod if right != 0 => Some(left % right),
                BinOp::BitAnd => Some(left & right),
                BinOp::BitOr => Some(left | right),
                BinOp::BitXor => Some(left ^ right),
                _ => None,
            }
        }
        _ => None,
    }
}

fn walk_procedure_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut interface_name: Option<String> = None;
    let mut declarations = Vec::new();

    for child in pair.into_inner().filter(|p| meaningful(p)) {
        match child.as_rule() {
            Rule::identifier if interface_name.is_none() => {
                interface_name = Some(child.as_str().to_string());
            }
            Rule::proc_decl_item => {
                let mut item_name: Option<String> = None;
                let mut init: Option<Expression> = None;
                for item_child in child.into_inner().filter(|p| meaningful(p)) {
                    match item_child.as_rule() {
                        Rule::identifier if item_name.is_none() => {
                            item_name = Some(item_child.as_str().to_string());
                        }
                        Rule::identifier => {
                            init = Some(Expression::new(ExprKind::FuncRef(
                                item_child.as_str().to_string(),
                            )));
                        }
                        _ => {
                            let text = item_child.as_str().trim();
                            if text.eq_ignore_ascii_case("null()") {
                                init = Some(Expression::null());
                            }
                        }
                    }
                }

                if let Some(name) = item_name {
                    let type_hint = interface_name
                        .as_ref()
                        .map(|iface| format!("procedure({iface})"))
                        .or_else(|| Some("procedure".to_string()));
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: type_hint.map(Into::into),
                        init,
                        array_bounds: None,
                        with_events: false,
                    });
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Dim,
    }))
}

fn parse_fortran_visibility(text: &str) -> Option<Visibility> {
    match text.trim().to_ascii_lowercase().as_str() {
        "public" => Some(Visibility::Public),
        "private" => Some(Visibility::Private),
        "protected" => Some(Visibility::Protected),
        _ => None,
    }
}

fn apply_fortran_type_attribute(
    pair: Pair<Rule>,
    modifiers: &mut ClassModifiers,
    parents: &mut Vec<String>,
) {
    if let Some(visibility) = parse_fortran_visibility(pair.as_str()) {
        modifiers.visibility = visibility;
        return;
    }

    if pair.as_str().trim().eq_ignore_ascii_case("abstract") {
        modifiers.is_abstract = true;
        return;
    }

    for child in pair.into_inner().filter(|child| meaningful(child)) {
        if child.as_rule() == Rule::identifier {
            parents.push(child.as_str().to_string());
        }
    }
}

fn apply_fortran_type_bound_attribute(text: &str, modifiers: &mut Modifiers) {
    if let Some(visibility) = parse_fortran_visibility(text) {
        modifiers.visibility = visibility;
        return;
    }

    let attr = text.trim();
    if attr.eq_ignore_ascii_case("deferred") {
        modifiers.is_abstract = true;
    } else if attr.eq_ignore_ascii_case("non_overridable") {
        modifiers.is_not_overridable = true;
    } else if attr.eq_ignore_ascii_case("nopass") {
        // `procedure, nopass :: s` binds a procedure that takes NO receiver —
        // it is the type's static method. Recorded so the binder stops
        // demanding a dummy argument of the type, which a `nopass` procedure
        // by definition does not have.
        modifiers.is_static = true;
    }
}

fn walk_assign(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let value = walk_expr(parts.pop().ok_or("missing rhs")?)?;
    let mut target = Expression::new(ExprKind::Ident(parts[0].as_str().to_string()));
    for p in &parts[1..] {
        if p.as_rule() == Rule::member_or_index {
            let mut inner = p.clone().into_inner().filter(|m| meaningful(m));
            let first = inner.next();
            match first {
                Some(m) if matches!(m.as_rule(), Rule::identifier | Rule::designator_name) => {
                    target = Expression::new(ExprKind::Member {
                        object: Box::new(target),
                        field: m.as_str().to_string(),
                        null_safe: false,
                    });
                }
                Some(m) if m.as_rule() == Rule::argument_list => {
                    let mut indices = Vec::new();
                    for arg in m
                        .into_inner()
                        .filter(|item| item.as_rule() == Rule::argument)
                    {
                        let (_, index) = walk_argument_expr(arg)?;
                        indices.push(index);
                    }
                    for index in indices {
                        target = Expression::new(ExprKind::Index {
                            object: Box::new(target),
                            index: Box::new(index),
                            null_safe: false,
                        });
                    }
                }
                Some(m) => {
                    let idx = walk_expr(m)?;
                    target = Expression::new(ExprKind::Index {
                        object: Box::new(target),
                        index: Box::new(idx),
                        null_safe: false,
                    });
                }
                None => {}
            }
        }
    }
    Ok(Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    }))
}

/// `WHERE (mask) … ELSEWHERE … END WHERE` — assignment under an elementwise
/// mask.
///
/// Every assignment in a clause keeps the elements the mask excludes, which is
/// exactly `merge(rhs, lhs, mask)` — the same intrinsic Fortran gives the
/// programmer, so no new machinery. `ELSEWHERE` narrows the mask by the
/// negation of every clause before it; the negation and the conjunction are
/// elementwise for the same reason the comparison that built the mask was.
///
/// The grammar has parsed this since the construct was added; the walker had
/// no arm for it, so the body ran UNMASKED — `where (v >= 2) r = v * 10`
/// assigned to every element.
fn walk_where(pair: Pair<Rule>) -> Result<Statement, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut clauses: Vec<(Option<Expression>, Vec<Statement>)> = Vec::new();
    let mut current_mask: Option<Option<Expression>> = None;
    let mut current_body: Vec<Statement> = Vec::new();
    let mut seen_first_mask = false;

    for p in inner {
        if p.as_rule() == Rule::kw_elsewhere {
            clauses.push((current_mask.take().unwrap_or(None), std::mem::take(&mut current_body)));
            current_mask = Some(None);
            continue;
        }
        if is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression {
            // The first expression is the WHERE mask; a later one belongs to
            // the `ELSEWHERE (mask)` clause just opened.
            if !seen_first_mask {
                seen_first_mask = true;
                current_mask = Some(Some(walk_expr(p)?));
            } else if matches!(current_mask, Some(None)) {
                current_mask = Some(Some(walk_expr(p)?));
            }
            continue;
        }
        match p.as_rule() {
            Rule::assignment_statement => {
                if let Some(st) = walk_stmt(p)? {
                    current_body.push(st);
                }
            }
            Rule::statement_line | Rule::line => {
                for s in p.into_inner().filter(|q| meaningful(q)) {
                    if let Some(st) = walk_stmt(s)? {
                        current_body.push(st);
                    }
                }
            }
            _ => {}
        }
    }
    clauses.push((current_mask.take().unwrap_or(None), current_body));

    let mut statements = Vec::new();
    let mut preceding: Vec<Expression> = Vec::new();
    for (mask, body) in clauses {
        let effective = fortran_where_effective_mask(mask.clone(), &preceding);
        if let Some(mask) = mask {
            preceding.push(mask);
        }
        for stmt in body {
            statements.push(fortran_mask_assignment(stmt, &effective));
        }
    }
    Ok(Statement::new(StmtKind::Block(statements)))
}

/// The mask a clause actually assigns under: its own, minus everything the
/// clauses before it already claimed.
/// Built as an explicit map rather than as `.not. m1 .and. m2` over whole
/// arrays: the elementwise repair fires on operands it can recognize as arrays
/// BY NAME, and a clause mask is usually an expression (`v < 0`). Left to the
/// repair, `.not. (v < 0)` stayed a scalar negation of an array — which is
/// `false`, so the ELSEWHERE clause assigned nothing at all.
fn fortran_where_effective_mask(
    mask: Option<Expression>,
    preceding: &[Expression],
) -> Option<Expression> {
    // The FIRST clause assigns under its own mask untouched. Mapping over it
    // would be pointless work and, for `where (.false.)`, actively wrong: the
    // mask is a SCALAR, and a map over a scalar yields undefined. `merge`
    // already asks at run time whether a mask is an array.
    if preceding.is_empty() {
        return mask;
    }

    let item_name = "__fortran_where_item";
    let index_name = "__fortran_where_index";
    let not_true = |expr: Expression| {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(fortran_expr_is_true(expr)),
        })
    };

    // A bare ELSEWHERE has no mask of its own, so the FIRST earlier clause
    // supplies both the shape to walk and the first exclusion.
    let (driver, mut condition, already_excluded) = match mask {
        Some(mask) => (
            mask,
            fortran_expr_is_true(Expression::ident(item_name)),
            0,
        ),
        None => (
            preceding.first()?.clone(),
            not_true(Expression::ident(item_name)),
            1,
        ),
    };

    for earlier in preceding.iter().skip(already_excluded) {
        condition = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(condition),
            right: Box::new(not_true(Expression::new(ExprKind::Index {
                object: Box::new(earlier.clone()),
                index: Box::new(Expression::ident(index_name)),
                null_safe: false,
            }))),
        });
    }

    Some(build_fortran_array_map(
        driver,
        condition,
        true,
        item_name,
        index_name,
    ))
}

/// `lhs = rhs` under `mask` is `lhs = merge(rhs, lhs, mask)`.
fn fortran_mask_assignment(stmt: Statement, mask: &Option<Expression>) -> Statement {
    let Some(mask) = mask else {
        return stmt;
    };
    let StmtKind::Assign {
        targets,
        value,
        by_ref,
    } = stmt.kind
    else {
        return stmt;
    };
    let Some(target) = targets.first().cloned() else {
        return Statement::new(StmtKind::Assign {
            targets,
            value,
            by_ref,
        });
    };
    // Built directly rather than as a `merge(...)` CALL: the intrinsic fold
    // runs while walking a call expression, and a node synthesized here never
    // passes through it — the call would reach the runtime as a function that
    // does not exist.
    let merged = fortran_merge_node(value, target, mask.clone());
    Statement::new(StmtKind::Assign {
        targets,
        value: merged,
        by_ref,
    })
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
            Rule::expression
            | Rule::logical_equiv
            | Rule::logical_or
            | Rule::logical_and
            | Rule::logical_not
            | Rule::comparison
            | Rule::addition
            | Rule::multiplication
            | Rule::power
            | Rule::concat
            | Rule::unary
            | Rule::primary_expr => {
                if cond.is_none() {
                    cond = Some(walk_expr(p)?);
                }
            }
            Rule::statement_line => {
                for s in p.into_inner().filter(|p| meaningful(p)) {
                    if let Some(st) = walk_stmt(s)? {
                        then_body.push(st);
                    }
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
                            if let Some(st) = walk_stmt(s)? {
                                eb.push(st);
                            }
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
                            if let Some(st) = walk_stmt(s)? {
                                eb.push(st);
                            }
                        }
                    }
                }
                else_body = Some(eb);
            }
            // Single-line if body. The list must cover every alternative of
            // the grammar's `single_statement`, because anything missing is
            // dropped in silence — `if (x == 0) goto 10` ran the statement it
            // was meant to jump over.
            Rule::print_statement
            | Rule::write_statement
            | Rule::call_statement
            | Rule::assignment_statement
            | Rule::return_statement
            | Rule::cycle_statement
            | Rule::exit_statement
            | Rule::stop_statement
            | Rule::error_stop_statement
            | Rule::allocate_statement
            | Rule::deallocate_statement
            | Rule::goto_statement
            | Rule::continue_statement
            | Rule::sync_statement
            | Rule::expression_statement => {
                if let Some(st) = walk_stmt(p)? {
                    then_body.push(st);
                }
            }
            _ => {} // skip keywords (kw_if, kw_then, kw_end, etc.)
        }
    }
    Ok(Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body,
        elifs,
        else_body,
    }))
}

fn walk_do(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // Collect: identifier, expression, expression [, expression], statement_line*
    let mut var = String::new();
    let mut exprs = Vec::new();
    let mut body_parts = Vec::new();
    let mut body = Vec::new();
    let mut label = None;
    let mut closed = false;
    for p in parts {
        match p.as_rule() {
            // `outer: do i = 1, n`. Its own grammar rule, so it can never be
            // mistaken for the loop variable or a bound.
            Rule::loop_label => {
                label = Some(p.as_str().to_string());
            }
            // Everything past `end do` is the CLOSING NAME, never part of the
            // header. `identifier` is an expr rule, so the `is_expr_rule` arm
            // below used to swallow it into `exprs`, where it landed in the
            // third slot and became the loop STEP: `end do outer` compiled to
            // `i = i + outer`, an undefined name, so every named loop ran
            // exactly one iteration. The dead `Rule::identifier => {}` arm that
            // used to sit after `is_expr_rule` was meant to catch this and
            // could never be reached.
            Rule::kw_end_do => {
                closed = true;
            }
            Rule::identifier if var.is_empty() && !closed => {
                var = p.as_str().to_string();
            }
            Rule::statement_line => {
                body_parts.push(p);
            }
            Rule::inline_statement_list => {
                body.extend(walk_inline_statement_list(p)?);
            }
            _ if is_expr_rule(p.as_rule()) && !closed => {
                exprs.push(p);
            }
            _ => {} // skip kw_do, the closing name, etc.
        }
    }
    body.extend(walk_body(body_parts.into_iter())?);
    if var.is_empty() {
        return Ok(label_fortran_loop(
            label,
            Statement::new(StmtKind::While {
                cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
                body,
                else_body: None,
            }),
        ));
    }

    let start = if !exprs.is_empty() {
        walk_expr(exprs.remove(0))?
    } else {
        Expression::new(ExprKind::Lit(Literal::Int(0)))
    };
    let end_e = if !exprs.is_empty() {
        walk_expr(exprs.remove(0))?
    } else {
        Expression::new(ExprKind::Lit(Literal::Int(0)))
    };
    let step_expr = if !exprs.is_empty() {
        Some(walk_expr(exprs.remove(0))?)
    } else {
        None
    };
    let init = Some(Box::new(Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Ident(var.clone()))],
        value: start,
        by_ref: false,
    })));
    let sv = step_expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
    // Fortran FREEZES the trip count at loop entry: `do i = 1, bound` with the
    // body assigning `bound = 10` still runs five times. The bound was left as
    // a live expression in the condition, so mutating it EXTENDED the loop —
    // `do_bound_expression_freeze` summed 1..10 instead of 1..5.
    //
    // The limit and the step are therefore evaluated ONCE, into locals named
    // after the loop variable so nested loops never share one. A literal cannot
    // change under anyone, so it stays inline and the common loop is untouched.
    let mut hoisted = Vec::new();
    let end_e = fortran_hoist_loop_bound(end_e, &format!("__fortran_do_limit_{var}"), &mut hoisted);
    let sv = fortran_hoist_loop_bound(sv, &format!("__fortran_do_step_{var}"), &mut hoisted);
    // A NEGATIVE step counts DOWN: `do i = 10, 1, -1` runs ten times, not zero.
    // The condition was hard-wired to `i <= end`, so every countdown loop was
    // skipped entirely.
    let cond = Some(match fortran_step_is_negative(&sv) {
        // The sign is settled at compile time — which covers every literal
        // step, the overwhelmingly common spelling — so the test stays a
        // single comparison.
        Some(negative) => Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
            op: if negative { BinOp::GtEq } else { BinOp::LtEq },
            right: Box::new(end_e),
        }),
        // `do i = a, b, s` — the direction is only known at RUNTIME.
        // `(end - i) * step >= 0` is the same test for both: a positive step
        // reduces to `i <= end`, a negative one to `i >= end`.
        None => Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::new(ExprKind::Binary {
                    left: Box::new(end_e),
                    op: BinOp::Sub,
                    right: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
                })),
                op: BinOp::Mul,
                right: Box::new(sv.clone()),
            })),
            op: BinOp::GtEq,
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
        }),
    });
    // i = i + step as an Assign expression
    let update = Some(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
        value: Box::new(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Ident(var))),
            op: BinOp::Add,
            right: Box::new(sv),
        })),
    }));
    let loop_stmt = label_fortran_loop(
        label,
        Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body,
        }),
    );
    if hoisted.is_empty() {
        return Ok(loop_stmt);
    }
    hoisted.push(loop_stmt);
    Ok(Statement::new(StmtKind::Block(hoisted)))
}

/// Evaluate a DO bound ONCE, unless it is a literal that nothing can change.
///
/// Returns the expression to use in the loop, pushing the evaluation onto
/// `hoisted` when one is needed.
fn fortran_hoist_loop_bound(
    bound: Expression,
    name: &str,
    hoisted: &mut Vec<Statement>,
) -> Expression {
    if matches!(bound.kind, ExprKind::Lit(_)) {
        return bound;
    }
    hoisted.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(name)],
        value: bound,
        by_ref: false,
    }));
    Expression::ident(name)
}

/// Whether a DO step counts DOWN, when that is decidable at compile time.
///
/// `None` means the sign is a runtime question — a variable step — and the
/// caller has to emit a direction-agnostic condition rather than guess.
fn fortran_step_is_negative(step: &Expression) -> Option<bool> {
    match &step.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value < 0),
        ExprKind::Lit(Literal::Float(value)) => Some(*value < 0.0),
        // `-1` reaches here as a negation of the literal `1`, which is how
        // the grammar spells it — checking only `Lit` would miss every
        // countdown loop in the corpus.
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => fortran_step_is_negative(expr).map(|negative| !negative),
        ExprKind::Unary {
            op: UnaryOp::Pos,
            expr,
        } => fortran_step_is_negative(expr),
        _ => None,
    }
}

/// The construct name on a `cycle` / `exit`, if the source gave one.
fn fortran_loop_target_name(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|inner| inner.as_rule() == Rule::identifier)
        .map(|inner| inner.as_str().to_string())
}

/// Wrap a loop in `StmtKind::Labeled` when the source named it.
///
/// The construct name is what `cycle <name>` / `exit <name>` target, and the
/// compiler already resolves `ContinueTarget::Label` / `BreakTarget::Label`
/// against this wrapper — the walker simply never produced one, so a named
/// loop was indistinguishable from an unnamed one and every multi-level
/// `cycle`/`exit` hit the innermost loop instead.
fn label_fortran_loop(label: Option<String>, loop_stmt: Statement) -> Statement {
    match label {
        Some(label) => Statement::new(StmtKind::Labeled {
            label,
            body: Box::new(loop_stmt),
        }),
        None => loop_stmt,
    }
}

/// Fortran 2008 BLOCK — a nested scope with its own declarations whose
/// executable statements run in place.
///
/// The grammar has had the construct all along, but `walk_stmt_inner` had no
/// arm for it, so it fell to `_ => Ok(None)` and the ENTIRE body between
/// `block` and `end block` was dropped without a diagnostic. That is why the
/// suite reported "0 lines, wanted N" rather than a wrong answer: the
/// statements never existed.
///
/// A named block is what `exit <name>` targets, exactly like a named loop, so
/// it wraps in the same `Labeled` node.
fn walk_block_construct(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut label = None;
    let mut body = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::loop_label if label.is_none() => label = Some(p.as_str().to_string()),
            Rule::statement_line => {
                for s in p.into_inner().filter(|q| meaningful(q)) {
                    if let Some(st) = walk_stmt(s)? {
                        body.push(st);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(label_fortran_loop(
        label,
        Statement::new(StmtKind::Block(body)),
    ))
}

/// The call a `goto` becomes until the dispatch pass rewrites it.
const FORTRAN_GOTO_MARKER: &str = "__fortran_goto";
/// Prefix of the `Labeled` wrapper a numbered statement gets.
const FORTRAN_LABEL_PREFIX: &str = "__fortran_label_";
/// Name of the loop the dispatch pass jumps back into.
const FORTRAN_DISPATCH_LABEL: &str = "__fortran_dispatch";
/// The variable holding which segment runs next.
const FORTRAN_DISPATCH_STATE: &str = "__fortran_state";

/// One `statement_line`, keeping the leading numeric label.
///
/// The label is a GOTO target, so it cannot be dropped the way the previous
/// per-statement walk dropped it. An unlabelled line still flattens to its
/// statements — wrapping every line would put a `Block` scope around
/// declarations that the surrounding passes need to see.
fn walk_statement_line_stmts(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut label = None;
    let mut stmts = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        if p.as_rule() == Rule::statement_label {
            if label.is_none() {
                label = Some(p.as_str().to_string());
            }
            continue;
        }
        if let Some(st) = walk_stmt(p)? {
            stmts.push(st);
        }
    }
    let Some(label) = label else {
        return Ok(stmts);
    };
    let body = match stmts.len() {
        // `10 format(...)` and friends walk to nothing, but the LABEL is still
        // a jump target and the line still has to exist for it to land on.
        0 => Statement::new(StmtKind::Block(Vec::new())),
        1 => stmts.remove(0),
        _ => Statement::new(StmtKind::Block(stmts)),
    };
    Ok(vec![Statement::new(StmtKind::Labeled {
        label: format!("{FORTRAN_LABEL_PREFIX}{label}"),
        body: Box::new(body),
    })])
}

fn fortran_first_statement_label(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|p| p.as_rule() == Rule::statement_label)
        .map(|p| p.as_str().to_string())
}

fn fortran_goto_marker_statement(label: String) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(FORTRAN_GOTO_MARKER)),
        args: vec![Argument::positional(Expression::string(&label))],
        optional: false,
    })))
}

/// `GOTO (10, 20, 30) i` — the index selects the label, 1-based, and an index
/// outside the list falls through to the next statement.
fn walk_computed_goto(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut labels = Vec::new();
    let mut selector = None;
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::statement_label => labels.push(p.as_str().to_string()),
            Rule::expression if selector.is_none() => selector = Some(walk_expr(p)?),
            _ => {}
        }
    }
    let selector = selector.ok_or("missing computed goto selector")?;
    let mut elifs = Vec::new();
    let mut labels = labels.into_iter().enumerate();
    let (_, first) = labels.next().ok_or("empty computed goto label list")?;
    for (index, label) in labels {
        elifs.push((
            fortran_index_equals(&selector, index as i64 + 1),
            vec![fortran_goto_marker_statement(label)],
        ));
    }
    Ok(Statement::new(StmtKind::If {
        cond: fortran_index_equals(&selector, 1),
        then_body: vec![fortran_goto_marker_statement(first)],
        elifs,
        else_body: None,
    }))
}

fn fortran_index_equals(selector: &Expression, value: i64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(selector.clone()),
        right: Box::new(Expression::int(value)),
    })
}

/// Arithmetic IF — `if (e) lneg, lzero, lpos`, branching on the SIGN of `e`.
fn walk_arithmetic_if(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut cond = None;
    let mut labels = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::statement_label => labels.push(p.as_str().to_string()),
            _ if is_expr_rule(p.as_rule()) && cond.is_none() => cond = Some(walk_expr(p)?),
            _ => {}
        }
    }
    let cond = cond.ok_or("missing arithmetic if expression")?;
    if labels.len() != 3 {
        return Err("arithmetic if needs three labels".to_string());
    }
    let compare = |op: BinOp| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(cond.clone()),
            right: Box::new(Expression::int(0)),
        })
    };
    Ok(Statement::new(StmtKind::If {
        cond: compare(BinOp::Lt),
        then_body: vec![fortran_goto_marker_statement(labels[0].clone())],
        elifs: vec![(
            compare(BinOp::Eq),
            vec![fortran_goto_marker_statement(labels[1].clone())],
        )],
        else_body: Some(vec![fortran_goto_marker_statement(labels[2].clone())]),
    }))
}

/// Turn a body that uses GOTO into a dispatch loop.
///
/// GOTO and its numbered targets were parsed and then dropped — `goto_basic`
/// printed the value the jump was supposed to skip. Both halves are markers by
/// the time this runs: a target is a `Labeled` wrapper named
/// `__fortran_label_<n>`, a jump is a call to `__fortran_goto("<n>")`.
///
/// The body splits at each targeted label into segments, and the segments
/// become the arms of
///
/// ```text
/// __fortran_dispatch: while (true) { if (state == 0) { … } else if … else break }
/// ```
///
/// so a jump is `state = <segment>; continue __fortran_dispatch`. One flat loop
/// is enough for the whole corpus: targets sit at procedure top level, while
/// jumps come from anywhere, and a LABELLED continue reaches the loop from
/// inside a nested `do` or `if` just as well.
fn lower_fortran_goto_dispatch(body: &mut Vec<Statement>) -> Result<(), String> {
    let mut next_id = 0;
    for statement in body.iter_mut() {
        lower_fortran_goto_dispatch_in_statement(statement)?;
    }
    restructure_fortran_goto_body(body, &mut next_id);
    // Anything left names a label that does not exist in its procedure. gfortran
    // rejects that at compile time, and leaving the marker in would surface as
    // "undefined is not callable" at run time instead.
    let mut unresolved = HashSet::new();
    for statement in body.iter() {
        collect_fortran_goto_targets(statement, &mut unresolved);
    }
    if let Some(label) = unresolved.iter().min() {
        return Err(format!("goto references undefined statement label {label}"));
    }
    Ok(())
}

fn lower_fortran_goto_dispatch_in_statement(statement: &mut Statement) -> Result<(), String> {
    match &mut statement.kind {
        StmtKind::FunctionDecl { body, .. } => lower_fortran_goto_dispatch(body)?,
        StmtKind::ModuleDecl { members, .. }
        | StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. } => {
            for member in members.iter_mut() {
                if let ClassMember::Method(method) = member {
                    lower_fortran_goto_dispatch_in_statement(method)?;
                }
            }
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts.iter_mut() {
                lower_fortran_goto_dispatch_in_statement(stmt)?;
            }
        }
        _ => {}
    }
    Ok(())
}

/// Restructure one statement list, after every list nested inside it.
///
/// Bottom-up, because a label can sit inside a loop — `100 continue` closing a
/// `do` body is the F77 spelling of `cycle`. The inner list handles the labels
/// it owns; a jump aimed further out is left as a marker and the enclosing list
/// picks it up, reaching back through the inner dispatch with a LABELLED
/// continue.
fn restructure_fortran_goto_body(body: &mut Vec<Statement>, next_id: &mut usize) {
    for statement in body.iter_mut() {
        restructure_fortran_goto_in_statement(statement, next_id);
    }
    let dispatch_label = format!("{FORTRAN_DISPATCH_LABEL}{}", *next_id);
    let dispatch_state = format!("{FORTRAN_DISPATCH_STATE}{}", *next_id);
    let mut targets = HashSet::new();
    for statement in body.iter() {
        collect_fortran_goto_targets(statement, &mut targets);
    }
    let boundaries: Vec<(usize, String)> = body
        .iter()
        .enumerate()
        .filter_map(|(index, statement)| {
            let StmtKind::Labeled { label, .. } = &statement.kind else {
                return None;
            };
            let number = label.strip_prefix(FORTRAN_LABEL_PREFIX)?;
            targets
                .contains(number)
                .then(|| (index, number.to_string()))
        })
        .collect();
    if boundaries.is_empty() {
        unwrap_fortran_label_markers(body);
        return;
    }
    *next_id += 1;

    let mut segment_of = HashMap::new();
    for (segment, (_, number)) in boundaries.iter().enumerate() {
        segment_of.insert(number.clone(), segment + 1);
    }

    let mut segments: Vec<Vec<Statement>> = Vec::with_capacity(boundaries.len() + 1);
    let mut rest = std::mem::take(body);
    for (index, _) in boundaries.iter().rev() {
        segments.push(rest.split_off(*index));
    }
    segments.push(rest);
    segments.reverse();

    // Declarations have to outlive the arm they were written in: every one of
    // them sits in segment 0, and every later segment refers to them. An arm is
    // a scope, so they are hoisted ahead of the loop and only their
    // initialisers stay behind, in place, as assignments.
    let mut prelude = Vec::new();
    for segment in segments.iter_mut() {
        unwrap_fortran_label_markers(segment);
        hoist_fortran_declarations(segment, &mut prelude);
    }
    for (segment, statements) in segments.iter_mut().enumerate() {
        for statement in statements.iter_mut() {
            rewrite_fortran_goto_markers(statement, &segment_of, &dispatch_label, &dispatch_state);
        }
        statements.push(fortran_dispatch_state_assignment(&dispatch_state, segment + 1));
        statements.push(Statement::new(StmtKind::Continue(ContinueTarget::Label(
            dispatch_label.clone(),
        ))));
    }

    let mut arms = segments.into_iter().enumerate();
    let (_, first) = arms.next().expect("segments is never empty");
    let dispatch = Statement::new(StmtKind::If {
        cond: fortran_dispatch_state_equals(&dispatch_state, 0),
        then_body: first,
        elifs: arms
            .map(|(segment, statements)| {
                (
                    fortran_dispatch_state_equals(&dispatch_state, segment as i64),
                    statements,
                )
            })
            .collect(),
        else_body: Some(vec![Statement::new(StmtKind::Break(BreakTarget::Label(
            dispatch_label.clone(),
        )))]),
    });

    *body = prelude;
    body.push(fortran_dispatch_state_assignment(&dispatch_state, 0));
    body.push(Statement::new(StmtKind::Labeled {
        label: dispatch_label,
        body: Box::new(Statement::new(StmtKind::While {
            cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
            body: vec![dispatch],
            else_body: None,
        })),
    }));
}

/// Each nested list gets its OWN dispatch loop, so the loop name and the state
/// variable are numbered — an inner `continue` must not reach the outer loop.
fn restructure_fortran_goto_in_statement(statement: &mut Statement, next_id: &mut usize) {
    if let StmtKind::Labeled { body, .. } = &mut statement.kind {
        restructure_fortran_goto_in_statement(body, next_id);
        return;
    }
    for_each_fortran_nested_vec_mut(&mut statement.kind, &mut |stmts| {
        restructure_fortran_goto_body(stmts, next_id)
    });
}

fn fortran_dispatch_state_equals(state: &str, segment: i64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::ident(state)),
        right: Box::new(Expression::int(segment)),
    })
}

fn fortran_dispatch_state_assignment(state: &str, segment: usize) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(state)],
        value: Expression::int(segment as i64),
        by_ref: false,
    })
}

/// Move declarations ahead of the dispatch loop, leaving each initialiser
/// behind as an assignment so it still runs where it was written.
fn hoist_fortran_declarations(segment: &mut Vec<Statement>, prelude: &mut Vec<Statement>) {
    let mut kept = Vec::with_capacity(segment.len());
    for statement in std::mem::take(segment) {
        match statement.kind {
            StmtKind::VarDecl { declarations, kind } => {
                let mut hoisted = Vec::with_capacity(declarations.len());
                for mut declaration in declarations {
                    if let Some(init) = declaration.init.take() {
                        if let BindingPattern::Ident(name) = &declaration.pattern {
                            kept.push(Statement::new(StmtKind::Assign {
                                targets: vec![Expression::ident(name)],
                                value: init,
                                by_ref: false,
                            }));
                        } else {
                            declaration.init = Some(init);
                        }
                    }
                    hoisted.push(declaration);
                }
                prelude.push(Statement::new(StmtKind::VarDecl {
                    declarations: hoisted,
                    kind,
                }));
            }
            StmtKind::FunctionDecl { .. }
            | StmtKind::ClassDecl { .. }
            | StmtKind::StructDecl { .. }
            | StmtKind::ModuleDecl { .. }
            | StmtKind::InterfaceDecl { .. } => prelude.push(Statement::new(statement.kind)),
            _ => kept.push(statement),
        }
    }
    *segment = kept;
}

fn collect_fortran_goto_targets(statement: &Statement, out: &mut HashSet<String>) {
    if let StmtKind::Expr(expr) = &statement.kind {
        if let Some(label) = fortran_goto_marker_label(expr) {
            out.insert(label);
        }
    }
    for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
        for child in stmts {
            collect_fortran_goto_targets(child, out);
        }
    });
}

/// Every nested statement list a jump can be written inside.
///
/// Procedure and class bodies are deliberately absent: a GOTO cannot leave the
/// procedure it is written in, so their labels belong to their own dispatch.
fn for_each_fortran_nested_body(kind: &StmtKind, f: &mut dyn FnMut(&[Statement])) {
    match kind {
        StmtKind::Block(stmts)
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::For { body: stmts, .. }
        | StmtKind::ForIn { body: stmts, .. }
        | StmtKind::With { body: stmts, .. }
        | StmtKind::Using { body: stmts, .. }
        | StmtKind::Lock { body: stmts, .. } => f(stmts),
        StmtKind::Labeled { body, .. } => f(std::slice::from_ref(body)),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            f(then_body);
            for (_, elif_body) in elifs {
                f(elif_body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                f(&case.body);
            }
            if let Some(default) = default {
                f(default);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            f(body);
            for catch in catches {
                f(&catch.body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
            if let Some(finally) = finally {
                f(finally);
            }
        }
        _ => {}
    }
}

/// Same shape as [`for_each_fortran_nested_body_mut`], but handing over the
/// whole `Vec` so a pass can REPLACE the list. `Labeled` is absent because its
/// body is a single statement, not a list — callers recurse into it directly.
fn for_each_fortran_nested_vec_mut(kind: &mut StmtKind, f: &mut dyn FnMut(&mut Vec<Statement>)) {
    match kind {
        StmtKind::Block(stmts)
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::For { body: stmts, .. }
        | StmtKind::ForIn { body: stmts, .. }
        | StmtKind::With { body: stmts, .. }
        | StmtKind::Using { body: stmts, .. }
        | StmtKind::Lock { body: stmts, .. } => f(stmts),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            f(then_body);
            for (_, elif_body) in elifs {
                f(elif_body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                f(&mut case.body);
            }
            if let Some(default) = default {
                f(default);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            f(body);
            for catch in catches {
                f(&mut catch.body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
            if let Some(finally) = finally {
                f(finally);
            }
        }
        _ => {}
    }
}

fn for_each_fortran_nested_body_mut(kind: &mut StmtKind, f: &mut dyn FnMut(&mut [Statement])) {
    match kind {
        StmtKind::Block(stmts)
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::For { body: stmts, .. }
        | StmtKind::ForIn { body: stmts, .. }
        | StmtKind::With { body: stmts, .. }
        | StmtKind::Using { body: stmts, .. }
        | StmtKind::Lock { body: stmts, .. } => f(stmts),
        StmtKind::Labeled { body, .. } => f(std::slice::from_mut(body.as_mut())),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            f(then_body);
            for (_, elif_body) in elifs {
                f(elif_body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                f(&mut case.body);
            }
            if let Some(default) = default {
                f(default);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            f(body);
            for catch in catches {
                f(&mut catch.body);
            }
            if let Some(else_body) = else_body {
                f(else_body);
            }
            if let Some(finally) = finally {
                f(finally);
            }
        }
        _ => {}
    }
}

fn fortran_goto_marker_label(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name != FORTRAN_GOTO_MARKER {
        return None;
    }
    match &args.first()?.value.kind {
        ExprKind::Lit(Literal::Str(label)) => Some(label.clone()),
        _ => None,
    }
}

/// Replace every `__fortran_goto("n")` with the jump into the dispatch loop.
fn rewrite_fortran_goto_markers(
    statement: &mut Statement,
    segment_of: &HashMap<String, usize>,
    dispatch_label: &str,
    dispatch_state: &str,
) {
    if let StmtKind::Expr(expr) = &statement.kind {
        if let Some(label) = fortran_goto_marker_label(expr) {
            if let Some(segment) = segment_of.get(&label) {
                *statement = Statement::new(StmtKind::Block(vec![
                    fortran_dispatch_state_assignment(dispatch_state, *segment),
                    Statement::new(StmtKind::Continue(ContinueTarget::Label(
                        dispatch_label.to_string(),
                    ))),
                ]));
                return;
            }
        }
    }
    for_each_fortran_nested_body_mut(&mut statement.kind, &mut |stmts| {
        for child in stmts {
            rewrite_fortran_goto_markers(child, segment_of, dispatch_label, dispatch_state);
        }
    });
}

/// Drop the `Labeled` wrapper a numbered line carries once nothing jumps to it.
fn unwrap_fortran_label_markers(body: &mut [Statement]) {
    for statement in body.iter_mut() {
        if let StmtKind::Labeled { label, body: inner } = &statement.kind {
            // A NAMED construct (`outer: do`) is also `Labeled` and must keep
            // its wrapper — `exit outer` resolves against it.
            if label.starts_with(FORTRAN_LABEL_PREFIX) {
                *statement = (**inner).clone();
            }
        }
        for_each_fortran_nested_body_mut(&mut statement.kind, &mut unwrap_fortran_label_markers);
    }
}

/// The `iso_c_binding` named constants, as values.
///
/// `c_null_char` terminates a C string, and the terminator has no counterpart
/// where strings carry their own length — so it contributes NOTHING, and
/// `c_char_"name"//c_null_char` is just `"name"`. The two null pointers are
/// null.
fn fortran_iso_c_binding_constant(name: &str) -> Option<Expression> {
    match name.to_ascii_lowercase().as_str() {
        "c_null_char" => Some(Expression::string("")),
        "c_null_ptr" | "c_null_funptr" => Some(Expression::new(ExprKind::Lit(Literal::Null))),
        _ => None,
    }
}

/// Pass `iso_c_binding` handles to a `bind(c)` procedure BY REFERENCE.
///
/// A Fortran dummy without `value` is passed by reference, and a C out-parameter
/// such as `sqlite3_open`'s `ppDb` is written THROUGH that reference. The C ABI
/// spells a reference as the shared carray pointer
/// (`primitives::pointers::make_carray_ptr` — `{__ref_kind, __base, __idx}`),
/// which is exactly what `&db` produces in C and what the platform adapter
/// reads back. Passing the handle's VALUE instead left the callee writing into
/// a null, so every one of these programs trapped on the first call.
///
/// Only declared handles travel this way. Everything else — the filename, the
/// SQL text — is an input the callee reads directly, so wrapping it would break
/// the argument it is supposed to deliver.
fn lower_fortran_c_binding_handles(body: &mut Vec<Statement>) {
    let mut handles = HashSet::new();
    collect_fortran_c_handle_names(body, &mut handles);
    if handles.is_empty() {
        return;
    }
    rewrite_fortran_c_handle_calls(body, &handles);
}

fn collect_fortran_c_handle_names(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration
                    .type_hint
                    .as_deref()
                    .is_some_and(is_fortran_opaque_c_handle)
                {
                    out.insert(name.to_ascii_lowercase());
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_c_handle_names(stmts, out)
        });
    }
}

fn rewrite_fortran_c_handle_calls(body: &mut Vec<Statement>, handles: &HashSet<String>) {
    let mut out = Vec::with_capacity(body.len());
    for mut statement in std::mem::take(body) {
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut |stmts| {
            rewrite_fortran_c_handle_calls(stmts, handles)
        });
        let call = match &mut statement.kind {
            StmtKind::Assign { value, .. } => value,
            StmtKind::Expr(expr) => expr,
            _ => {
                out.push(statement);
                continue;
            }
        };
        let ExprKind::Call { args, .. } = &mut call.kind else {
            out.push(statement);
            continue;
        };
        // The cell is a one-element array so the callee's write lands somewhere
        // the caller can read back; the write-back after the call is what makes
        // the reference observable in Fortran.
        let mut cells = Vec::new();
        for arg in args.iter_mut() {
            let ExprKind::Ident(name) = &arg.value.kind else {
                continue;
            };
            let name = name.clone();
            if !handles.contains(&name.to_ascii_lowercase()) {
                continue;
            }
            let cell = format!("__fortran_cbind_{}", name.to_ascii_lowercase());
            arg.value = vybe_compiler::primitives::pointers::make_carray_ptr(
                Expression::ident(&cell),
                Expression::int(0),
            );
            cells.push((cell, name.clone()));
        }
        if cells.is_empty() {
            out.push(statement);
            continue;
        }
        let mut block = Vec::with_capacity(cells.len() * 2 + 1);
        for (cell, name) in &cells {
            block.push(fortran_data_local(
                cell,
                fortran_array_expr(vec![Expression::ident(name)]),
            ));
        }
        block.push(statement);
        for (cell, name) in &cells {
            block.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(name)],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::ident(cell)),
                    index: Box::new(Expression::int(0)),
                    null_safe: false,
                }),
                by_ref: false,
            }));
        }
        out.push(Statement::new(StmtKind::Block(block)));
    }
    *body = out;
}

/// `read(buf, …) n` where `buf` is a CHARACTER variable — an internal file
/// that no internal WRITE ever filled.
///
/// Internal I/O works today through records: `write(buf, …)` stores one under
/// the buffer's handle and `read(buf, …)` reads it back. A buffer that was
/// merely ASSIGNED — `character(len=6) :: b = '6'` — has no record, so the read
/// found nothing and every target came back undefined.
///
/// ⛔ Only buffers never written are rewritten. Replacing the record path
/// wholesale scored −30: the write-then-read pairs that already worked lost the
/// mechanism they depend on.
///
/// The text is the buffer's own value, so the values are its blank/comma
/// separated tokens, converted by each target's declared type.
fn lower_fortran_internal_reads(body: &mut Vec<Statement>) {
    let mut declared = HashMap::new();
    collect_fortran_declared_types(body, &mut declared);
    let mut written = HashSet::new();
    collect_fortran_written_buffers(body, &mut written);
    declared.retain(|name, _| !written.contains(name));
    if declared.is_empty() {
        return;
    }
    rewrite_fortran_internal_reads(body, &declared);
}

fn collect_fortran_declared_types(body: &[Statement], out: &mut HashMap<String, String>) {
    for statement in body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                let Some(hint) = declaration.type_hint.as_deref() else {
                    continue;
                };
                if declaration.array_bounds.is_none() {
                    out.insert(name.to_ascii_lowercase(), hint.to_string());
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_declared_types(stmts, out)
        });
    }
}

/// Buffers an internal WRITE fills — those keep the record path.
fn collect_fortran_written_buffers(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        let target = match &statement.kind {
            StmtKind::PrintFile { file_number, .. } | StmtKind::WriteFile { file_number, .. } => {
                Some(file_number)
            }
            _ => None,
        };
        if let Some(ExprKind::Ident(name)) = target.map(|expr| &expr.kind) {
            out.insert(name.to_ascii_lowercase());
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_written_buffers(stmts, out)
        });
    }
}

fn rewrite_fortran_internal_reads(body: &mut Vec<Statement>, declared: &HashMap<String, String>) {
    for statement in body.iter_mut() {
        // The marker is handled BEFORE descending. Recursing first let the
        // generic list-directed arm rewrite the very `InputFile` the marker
        // describes, and the format was then applied to nothing.
        // A marked read: the format rides in front of the `InputFile`. Matched
        // by SCANNING the block rather than by expecting an exact two-statement
        // shape — any pass that inserts between them would break a positional
        // match, and the read would then silently take the list-directed path.
        // The marker is consumed in every case, so none can survive to run.
        if let StmtKind::Block(stmts) = &mut statement.kind {
            if let Some(format_spec) = fortran_take_read_format_marker(stmts) {
                for inner in stmts.iter_mut() {
                    let Some((buffer, variables)) =
                        fortran_internal_read_target(inner, declared)
                    else {
                        continue;
                    };
                    *inner = build_fortran_formatted_internal_read(
                        buffer,
                        variables,
                        &format_spec,
                        declared,
                    );
                }
                continue;
            }
        }
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut |stmts| {
            rewrite_fortran_internal_reads(stmts, declared)
        });
        let StmtKind::InputFile {
            file_number,
            variables,
        } = &statement.kind
        else {
            continue;
        };
        let ExprKind::Ident(name) = &file_number.kind else {
            continue;
        };
        if !declared
            .get(&name.to_ascii_lowercase())
            .is_some_and(|hint| is_fortran_string_type_hint(hint))
            || variables.is_empty()
        {
            continue;
        }
        *statement = build_fortran_internal_read(file_number.clone(), variables.clone(), declared);
    }
}

fn build_fortran_internal_read(
    buffer: Expression,
    variables: Vec<Expression>,
    declared: &HashMap<String, String>,
) -> Statement {
    let tokens = "__fortran_iread_tokens";
    let mut body = vec![fortran_data_local(
        tokens,
        fortran_internal_read_tokens(buffer),
    )];
    for (index, target) in variables.into_iter().enumerate() {
        let token = Expression::new(ExprKind::Index {
            object: Box::new(Expression::ident(tokens)),
            index: Box::new(Expression::int(index as i64)),
            null_safe: false,
        });
        let hint = match &target.kind {
            ExprKind::Ident(name) => declared.get(&name.to_ascii_lowercase()).cloned(),
            _ => None,
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value: fortran_internal_read_value(hint.as_deref(), token),
            by_ref: false,
        }));
    }
    Statement::new(StmtKind::Block(body))
}

/// Remove the format marker from a block and return the format it carried.
fn fortran_take_read_format_marker(stmts: &mut Vec<Statement>) -> Option<String> {
    let mut found = None;
    stmts.retain(|statement| {
        let StmtKind::Expr(expr) = &statement.kind else {
            return true;
        };
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return true;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return true;
        };
        if name != FORTRAN_READ_FORMAT_MARKER {
            return true;
        }
        if let Some(ExprKind::Lit(Literal::Str(format_spec))) =
            args.first().map(|arg| &arg.value.kind)
        {
            found = Some(format_spec.clone());
        }
        false
    });
    found
}

/// The buffer and targets of a read whose unit is an internal file.
fn fortran_internal_read_target(
    read: &Statement,
    declared: &HashMap<String, String>,
) -> Option<(Expression, Vec<Expression>)> {
    let StmtKind::InputFile {
        file_number,
        variables,
    } = &read.kind
    else {
        return None;
    };
    let ExprKind::Ident(name) = &file_number.kind else {
        return None;
    };
    if variables.is_empty()
        || !declared
            .get(&name.to_ascii_lowercase())
            .is_some_and(|hint| is_fortran_string_type_hint(hint))
    {
        return None;
    }
    Some((file_number.clone(), variables.clone()))
}

/// `read(buf, '(I3,1X,I3)') a, b` — a FORMATTED internal read.
///
/// The mirror of the write side: the same `parse_fortran_format_chunks` walked
/// the other way, so repeat counts, `nX` skips and literals cost nothing extra
/// and the descriptor grammar is understood in ONE place. Each data descriptor
/// consumes its own width from the buffer at a running position; `A` keeps the
/// characters, everything else converts.
///
/// Falls back to the list-directed token split when the format has no fixed
/// widths to consume — an `I0`-style descriptor names no field.
fn build_fortran_formatted_internal_read(
    buffer: Expression,
    variables: Vec<Expression>,
    format_spec: &str,
    declared: &HashMap<String, String>,
) -> Statement {
    let Some(chunks) = parse_fortran_format_chunks(format_spec) else {
        return build_fortran_internal_read(buffer, variables, declared);
    };
    let mut position = 0usize;
    let mut fields = Vec::new();
    for chunk in &chunks {
        match chunk {
            FortranFormatChunk::Spaces(count) => position += count,
            FortranFormatChunk::Literal(text) => position += text.chars().count(),
            FortranFormatChunk::Newline => {}
            FortranFormatChunk::Data {
                descriptor,
                repeat,
                width,
                ..
            } => {
                let Some(width) = width else {
                    return build_fortran_internal_read(buffer, variables, declared);
                };
                for _ in 0..(*repeat).max(1) {
                    fields.push((descriptor.clone(), position, *width));
                    position += width;
                }
            }
        }
    }
    if fields.len() < variables.len() {
        return build_fortran_internal_read(buffer, variables, declared);
    }

    let mut body = Vec::new();
    for (target, (descriptor, start, width)) in variables.into_iter().zip(fields) {
        let field = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__fortran_substring")),
            args: vec![
                Argument::positional(buffer.clone()),
                Argument::positional(Expression::int(start as i64)),
                Argument::positional(Expression::int((start + width) as i64)),
            ],
            optional: false,
        });
        // `A` delivers the characters as written, blanks included; every other
        // descriptor names a value, and a value never keeps its padding.
        let value = if descriptor.eq_ignore_ascii_case("a") {
            field
        } else {
            let trimmed = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("adjustl")),
                args: vec![Argument::positional(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("trim")),
                    args: vec![Argument::positional(field)],
                    optional: false,
                }))],
                optional: false,
            });
            let hint = match &target.kind {
                ExprKind::Ident(name) => declared.get(&name.to_ascii_lowercase()).cloned(),
                _ => None,
            };
            fortran_internal_read_value(hint.as_deref(), trimmed)
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        }));
    }
    Statement::new(StmtKind::Block(body))
}

/// `buf` → its blank/comma separated tokens, empties dropped.
///
/// Commas become blanks first so one split serves both separators, and the
/// filter is what makes `'8, 9'` two tokens rather than three — a plain split
/// on `" "` leaves the gap between the comma and the digit as an empty string.
fn fortran_internal_read_tokens(buffer: Expression) -> Expression {
    // Spelled with Fortran's own intrinsics and the declared `[array_methods]`
    // entries, not raw JS member calls: a bare `.trim()` on a value the
    // compiler has not typed as a string compiles to a struct field read and
    // an invoke of undefined.
    let intrinsic = |name: &str, args: Vec<Expression>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false,
        })
    };
    let split = intrinsic(
        "str_split",
        vec![
            intrinsic("trim", vec![buffer]),
            Expression::string(" "),
        ],
    );
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(split),
            field: "filter".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__fortran_iread_token".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::StrictNotEq,
                left: Box::new(Expression::ident("__fortran_iread_token")),
                right: Box::new(Expression::string("")),
            }))),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    })
}

/// One token, converted for the target's DECLARED type. Static, because the
/// target has no value yet to inspect.
fn fortran_internal_read_value(declared: Option<&str>, token: Expression) -> Expression {
    let hint = declared.unwrap_or("").to_ascii_lowercase();
    if is_fortran_string_type_hint(&hint) {
        return token;
    }
    if hint.starts_with("logical") {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(token.clone()),
                right: Box::new(Expression::string("T")),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(token),
                right: Box::new(Expression::string(".true.")),
            })),
        });
    }
    // A token that is not a number leaves the target UNDEFINED in Fortran and
    // sets `iostat`; storing the NaN instead traps on the integer conversion.
    // `iostat` is not carried on `InputFile`, so the status cannot be reported
    // here — but a non-numeric token must not crash the program.
    let number = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__fortran_to_number")),
        args: vec![Argument::positional(token)],
        optional: false,
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(number.clone()),
            right: Box::new(number.clone()),
        })),
        then: Box::new(number),
        else_: Box::new(Expression::int(0)),
    })
}

/// `17 / 5` is `3` — INTEGER division truncates toward zero.
///
/// Division always produced a real, and only an assignment to an integer target
/// rounded it back, so `17 / 5` printed `3.4` and `(17 / 5) /= 3` was true.
/// The operands decide the operation in Fortran, not the destination.
///
/// Static on purpose: a literal is an integer, and a variable is one when it was
/// DECLARED integer. Anything the walker cannot prove integer keeps real
/// division, so a wrong guess can only leave today's behaviour, never invent
/// truncation where the operands are real.
fn lower_fortran_integer_division(body: &mut [Statement]) {
    let mut declared = HashSet::new();
    collect_fortran_declared_integers(body, &mut declared);
    for statement in body.iter_mut() {
        statement.walk_exprs_mut(&mut |expr| {
            let ExprKind::Binary {
                op: BinOp::Div,
                left,
                right,
            } = &expr.kind
            else {
                return;
            };
            if !fortran_expr_is_integer(left, &declared)
                || !fortran_expr_is_integer(right, &declared)
            {
                return;
            }
            if let ExprKind::Binary { op, .. } = &mut expr.kind {
                *op = BinOp::IDiv;
            }
        });
    }
}

fn collect_fortran_declared_integers(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("integer"))
                {
                    out.insert(name.to_ascii_lowercase());
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_declared_integers(stmts, out)
        });
    }
}

/// Whether an expression is provably an INTEGER — conservative, so `false`
/// means "not proven", never "is real".
fn fortran_expr_is_integer(expr: &Expression, declared: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => true,
        ExprKind::Ident(name) => declared.contains(&name.to_ascii_lowercase()),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => fortran_expr_is_integer(expr, declared),
        ExprKind::Binary {
            op: BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::IDiv | BinOp::Mod,
            left,
            right,
        } => fortran_expr_is_integer(left, declared) && fortran_expr_is_integer(right, declared),
        // The intrinsics whose result is an integer whatever went in.
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name) if matches!(
                name.to_ascii_lowercase().as_str(),
                "int" | "nint" | "size" | "len" | "len_trim" | "count" | "index" | "ubound"
                    | "lbound" | "iachar" | "ichar" | "mod" | "modulo" | "floor" | "ceiling"
            )
        ),
        _ => false,
    }
}

/// `interface assignment(=)` — DEFINED assignment.
///
/// `b = 42` where `b` is a derived type and the module declares
/// `interface assignment(=) ; module procedure int_to_box`, is a CALL, not a
/// store: `call int_to_box(b, 42)`. Nothing dispatched it, so the store ran
/// instead and the target kept the raw right-hand side.
///
/// The interface names the implementations; their signatures live on the
/// module's own procedures, so both are collected and matched by the DECLARED
/// types of the two dummies. A rewrite happens only when exactly one candidate
/// matches — an ambiguous set is left alone rather than guessed at.
fn lower_fortran_defined_assignment(body: &mut Vec<Statement>) {
    let mut names = HashSet::new();
    collect_fortran_defined_assignment_names(body, &mut names);
    if names.is_empty() {
        return;
    }
    let mut procs = HashMap::new();
    collect_fortran_procedure_params(body, &mut procs);
    let candidates: Vec<(String, String, String)> = names
        .iter()
        .filter_map(|name| {
            let params = procs.get(name)?;
            let dest = params.first()?.type_hint.as_deref()?.to_string();
            let src = params.get(1)?.type_hint.as_deref()?.to_string();
            Some((dest, src, name.clone()))
        })
        .collect();
    if candidates.is_empty() {
        return;
    }
    let mut declared = HashMap::new();
    collect_fortran_declared_types(body, &mut declared);
    rewrite_fortran_defined_assignment(body, &candidates, &declared);
}

fn collect_fortran_defined_assignment_names(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        if let StmtKind::InterfaceDecl { name, members, .. } = &statement.kind {
            if name.trim().eq_ignore_ascii_case("assignment(=)") {
                for member in members {
                    if let InterfaceMember::Method { name, .. } = member {
                        out.insert(name.to_ascii_lowercase());
                    }
                }
            }
        }
        if let StmtKind::ModuleDecl { members, .. } = &statement.kind {
            for member in members {
                if let ClassMember::NestedType(inner) = member {
                    collect_fortran_defined_assignment_names(std::slice::from_ref(inner), out);
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_defined_assignment_names(stmts, out)
        });
    }
}

fn collect_fortran_procedure_params(body: &[Statement], out: &mut HashMap<String, Vec<Param>>) {
    for statement in body {
        if let StmtKind::FunctionDecl { name, params, .. } = &statement.kind {
            out.insert(name.to_ascii_lowercase(), params.clone());
        }
        if let StmtKind::ModuleDecl { members, .. } | StmtKind::ClassDecl { members, .. } =
            &statement.kind
        {
            for member in members {
                match member {
                    ClassMember::Method(inner) | ClassMember::NestedType(inner) => {
                        collect_fortran_procedure_params(std::slice::from_ref(inner), out)
                    }
                    _ => {}
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_procedure_params(stmts, out)
        });
    }
}

/// Whether a value can be the source argument of a dummy declared `hint`.
fn fortran_value_matches_hint(
    value: &Expression,
    hint: &str,
    declared: &HashMap<String, String>,
) -> bool {
    let hint_lower = hint.trim().to_ascii_lowercase();
    if let Some(wanted) = parse_derived_type_name(hint) {
        let ExprKind::Ident(name) = &value.kind else {
            return false;
        };
        return declared
            .get(&name.to_ascii_lowercase())
            .and_then(|actual| parse_derived_type_name(actual))
            .is_some_and(|actual| actual.eq_ignore_ascii_case(&wanted));
    }
    let family = |h: &str| {
        if h.starts_with("character") {
            "character"
        } else if h.starts_with("integer") {
            "integer"
        } else if h.starts_with("real") || h.starts_with("double") {
            "real"
        } else if h.starts_with("logical") {
            "logical"
        } else {
            "?"
        }
    };
    let wanted = family(&hint_lower);
    match &value.kind {
        ExprKind::Lit(Literal::Int(_)) => wanted == "integer",
        ExprKind::Lit(Literal::Float(_)) => wanted == "real",
        ExprKind::Lit(Literal::Str(_)) => wanted == "character",
        ExprKind::Lit(Literal::Bool(_)) => wanted == "logical",
        ExprKind::Ident(name) => declared
            .get(&name.to_ascii_lowercase())
            .is_some_and(|actual| family(&actual.trim().to_ascii_lowercase()) == wanted),
        _ => false,
    }
}

fn rewrite_fortran_defined_assignment(
    body: &mut Vec<Statement>,
    candidates: &[(String, String, String)],
    declared: &HashMap<String, String>,
) {
    for statement in body.iter_mut() {
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut |stmts| {
            rewrite_fortran_defined_assignment(stmts, candidates, declared)
        });
        let StmtKind::Assign { targets, value, .. } = &statement.kind else {
            continue;
        };
        let [target] = targets.as_slice() else {
            continue;
        };
        let ExprKind::Ident(name) = &target.kind else {
            continue;
        };
        let Some(target_type) = declared
            .get(&name.to_ascii_lowercase())
            .and_then(|hint| parse_derived_type_name(hint))
        else {
            continue;
        };
        let mut matched = candidates.iter().filter(|(dest, src, _)| {
            parse_derived_type_name(dest)
                .is_some_and(|dest| dest.eq_ignore_ascii_case(&target_type))
                && fortran_value_matches_hint(value, src, declared)
        });
        let Some((_, _, proc)) = matched.next() else {
            continue;
        };
        if matched.next().is_some() {
            continue;
        }
        *statement = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(proc)),
            args: vec![
                Argument::positional(target.clone()),
                Argument::positional(value.clone()),
            ],
            optional: false,
        })));
    }
}

/// Names an `EXTERNAL` statement declares.
const FORTRAN_EXTERNAL_MARKER: &str = "__fortran_external";

/// `external f` — f is a procedure, so drop any variable declaration of it.
///
/// A Fortran program spells an external function's RESULT type with an ordinary
/// type declaration (`integer f`), which is not a variable. Emitted as one, the
/// name held `0` and shadowed the procedure, so `f()` reported "f64 is not
/// callable".
fn lower_fortran_external_declarations(body: &mut Vec<Statement>) {
    let mut names = HashSet::new();
    collect_fortran_external_names(body, &mut names);
    strip_fortran_external_markers(body, &names);
}

fn collect_fortran_external_names(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        // A procedure body carries its own `external` statements, and the
        // marker must be found there too — left behind it becomes a call to
        // an undefined name.
        if let StmtKind::FunctionDecl { body, .. } = &statement.kind {
            collect_fortran_external_names(body, out);
        }
        if let StmtKind::Expr(expr) = &statement.kind {
            if let ExprKind::Call { callee, args, .. } = &expr.kind {
                if matches!(&callee.kind, ExprKind::Ident(n) if n == FORTRAN_EXTERNAL_MARKER) {
                    for arg in args {
                        if let ExprKind::Lit(Literal::Str(name)) = &arg.value.kind {
                            out.insert(name.to_ascii_lowercase());
                        }
                    }
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_external_names(stmts, out)
        });
    }
}

fn strip_fortran_external_markers(body: &mut Vec<Statement>, names: &HashSet<String>) {
    for statement in body.iter_mut() {
        // ⛔ A DUMMY procedure argument is declared in the body like any other
        // name, but it is a parameter — dropping its declaration is not the
        // same fix and breaks the call. Its own scope excludes it.
        if let StmtKind::FunctionDecl { params, body, .. } = &mut statement.kind {
            let mut inner: HashSet<String> = names.clone();
            for param in params.iter() {
                inner.remove(&param.name.to_ascii_lowercase());
            }
            strip_fortran_external_markers(body, &inner);
        }
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut |stmts| {
            strip_fortran_external_markers(stmts, names)
        });
    }
    body.retain(|statement| {
        if let StmtKind::Expr(expr) = &statement.kind {
            if let ExprKind::Call { callee, .. } = &expr.kind {
                if matches!(&callee.kind, ExprKind::Ident(n) if n == FORTRAN_EXTERNAL_MARKER) {
                    return false;
                }
            }
        }
        true
    });
    if names.is_empty() {
        return;
    }
    for statement in body.iter_mut() {
        let StmtKind::VarDecl { declarations, .. } = &mut statement.kind else {
            continue;
        };
        declarations.retain(|declaration| {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                return true;
            };
            // Only a bare type declaration is the procedure's result type; one
            // with an initialiser or bounds is a genuine variable.
            !(names.contains(&name.to_ascii_lowercase())
                && declaration.init.is_none()
                && declaration.array_bounds.is_none())
        });
    }
    body.retain(|statement| !matches!(&statement.kind, StmtKind::VarDecl { declarations, .. } if declarations.is_empty()));
}

/// `interface operator(+)` — bind the derived type's operator to its SLOT.
///
/// flexclassplan §2b route 2: the walker already parses the form, so it
/// registers the binding directly instead of fabricating a pseudo-name. The
/// binding is keyed by `protocol_slot_key`, a stable slot ID — never by the
/// procedure's spelling — so `emit_rich_binop` finds it without any shared code
/// naming Fortran (§0.2, §0.4).
///
/// ⛔ `//` (concat) and `.myop.` (a defined operator) have NO core slot, and
/// `BinOp::Concat` never consults one. flexclassplan §2c-bis calls for a
/// `Concat` slot and an extension registry for custom operators, and names this
/// exact Fortran case as the motivation; neither exists yet, so those forms are
/// left alone rather than bent onto a slot that means something else.
fn lower_fortran_operator_slots(body: &mut Vec<Statement>) {
    let mut bindings = Vec::new();
    collect_fortran_operator_interfaces(body, &mut bindings);
    if bindings.is_empty() {
        return;
    }
    let mut procs = HashMap::new();
    collect_fortran_procedure_decls(body, &mut procs);
    let mut methods: HashMap<String, Vec<Statement>> = HashMap::new();
    for (symbol, impl_name) in bindings {
        let Some(decl) = procs.get(&impl_name) else {
            continue;
        };
        let StmtKind::FunctionDecl { params, .. } = &decl.kind else {
            continue;
        };
        let Some(receiver) = params.first() else {
            continue;
        };
        let Some(type_name) = receiver
            .type_hint
            .as_deref()
            .and_then(parse_derived_type_name)
        else {
            continue;
        };
        // Arity picks between the binary and unary reading of `-`.
        let Some(slot) = fortran_operator_slot(&symbol, params.len()) else {
            continue;
        };
        let StmtKind::FunctionDecl {
            params,
            return_type,
            body,
            is_sub,
            ..
        } = &decl.kind
        else {
            continue;
        };
        // Same shape the type-bound binder produces: the implementation's own
        // params and body, under the slot's key. The receiver is the first
        // dummy, which is the convention a Fortran type-bound procedure
        // already uses.
        methods
            .entry(type_name.to_ascii_lowercase())
            .or_default()
            .push(Statement::new(StmtKind::FunctionDecl {
                name: vybe_ast::protocol_slot_key(slot),
                params: params.clone(),
                return_type: return_type.clone(),
                body: body.clone(),
                modifiers: Modifiers::default(),
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: *is_sub,
            }));
    }
    if !methods.is_empty() {
        attach_fortran_slot_methods(body, &methods);
    }
}

/// The core slot a Fortran operator token names, if there is one.
fn fortran_operator_slot(symbol: &str, arity: usize) -> Option<vybe_ast::ProtocolSlot> {
    use vybe_ast::ProtocolSlot as S;
    Some(match (symbol, arity) {
        ("+", 1) => S::Pos,
        ("-", 1) => S::Neg,
        (".not.", 1) => S::Not,
        ("+", _) => S::Add,
        ("-", _) => S::Sub,
        ("*", _) => S::Mul,
        ("/", _) => S::Div,
        ("**", _) => S::Pow,
        ("==", _) | (".eq.", _) => S::Eq,
        ("/=", _) | (".ne.", _) => S::Ne,
        ("<", _) | (".lt.", _) => S::Lt,
        ("<=", _) | (".le.", _) => S::Le,
        (">", _) | (".gt.", _) => S::Gt,
        (">=", _) | (".ge.", _) => S::Ge,
        (".and.", _) => S::And,
        (".or.", _) => S::Or,
        _ => return None,
    })
}

fn collect_fortran_operator_interfaces(body: &[Statement], out: &mut Vec<(String, String)>) {
    for statement in body {
        if let StmtKind::InterfaceDecl { name, members, .. } = &statement.kind {
            let trimmed = name.trim().to_ascii_lowercase();
            if let Some(symbol) = trimmed
                .strip_prefix("operator(")
                .and_then(|rest| rest.strip_suffix(')'))
            {
                let symbol = symbol.trim().to_string();
                for member in members {
                    if let InterfaceMember::Method { name, .. } = member {
                        out.push((symbol.clone(), name.to_ascii_lowercase()));
                    }
                }
            }
        }
        if let StmtKind::ModuleDecl { members, .. } = &statement.kind {
            for member in members {
                if let ClassMember::NestedType(inner) | ClassMember::Method(inner) = member {
                    collect_fortran_operator_interfaces(std::slice::from_ref(inner), out);
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_operator_interfaces(stmts, out)
        });
    }
}

fn collect_fortran_procedure_decls(body: &[Statement], out: &mut HashMap<String, Statement>) {
    for statement in body {
        if let StmtKind::FunctionDecl { name, .. } = &statement.kind {
            out.insert(name.to_ascii_lowercase(), statement.clone());
        }
        if let StmtKind::ModuleDecl { members, .. } | StmtKind::ClassDecl { members, .. } =
            &statement.kind
        {
            for member in members {
                if let ClassMember::NestedType(inner) | ClassMember::Method(inner) = member {
                    collect_fortran_procedure_decls(std::slice::from_ref(inner), out);
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_procedure_decls(stmts, out)
        });
    }
}

fn attach_fortran_slot_methods(
    body: &mut [Statement],
    methods: &HashMap<String, Vec<Statement>>,
) {
    for statement in body.iter_mut() {
        if let StmtKind::ClassDecl { name, members, .. } = &mut statement.kind {
            if let Some(slots) = methods.get(&name.to_ascii_lowercase()) {
                for slot_method in slots {
                    members.push(ClassMember::Method(Box::new(slot_method.clone())));
                }
            }
        }
        if let StmtKind::ModuleDecl { members, .. } = &mut statement.kind {
            for member in members.iter_mut() {
                if let ClassMember::NestedType(inner) | ClassMember::Method(inner) = member {
                    let mut one = vec![(**inner).clone()];
                    attach_fortran_slot_methods(&mut one, methods);
                    if let Some(first) = one.into_iter().next() {
                        **inner = first;
                    }
                }
            }
        }
        for_each_fortran_nested_body_mut(&mut statement.kind, &mut |stmts| {
            attach_fortran_slot_methods(stmts, methods)
        });
    }
}

/// `DATA v /vals/` — the F77 initializer.
///
/// The grammar has carried the construct all along but `walk_stmt_inner` had no
/// arm, so every DATA statement fell to `_ => Ok(None)` and the variable kept
/// its default 0. Lowered here to ordinary assignments at the point of the
/// statement: Fortran requires DATA to precede the executable part, so "in
/// place" is already "before first use".
///
/// A subscripted target is built as a `Call`, which is exactly the shape
/// `walk_assign` produces for `a(1) = x` — the existing subscript repair then
/// turns it into a 1-based `Index` like any other.
fn walk_data_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    for set in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::data_set)
    {
        let mut targets = Vec::new();
        let mut values = Vec::new();
        let mut loop_form = None;
        for part in set.into_inner().filter(|p| meaningful(p)) {
            match part.as_rule() {
                Rule::data_var_list => {
                    for var in part.into_inner().filter(|p| p.as_rule() == Rule::data_var) {
                        if let Some(form) = collect_fortran_data_targets(var, &mut targets)? {
                            loop_form = Some(form);
                        }
                    }
                }
                Rule::data_value_list => {
                    for value in part
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::data_value)
                    {
                        collect_fortran_data_values(value, &mut values)?;
                    }
                }
                _ => {}
            }
        }
        if let Some(form) = loop_form {
            body.push(build_fortran_data_loop(form, values));
            continue;
        }
        // One bare name against several values is the whole-array form
        // (`data a /1, 2, 3/`); everything else pairs off one for one.
        if targets.len() == 1 && values.len() > 1 && matches!(targets[0].kind, ExprKind::Ident(_)) {
            body.push(fortran_data_assignment(
                targets.remove(0),
                fortran_array_expr(values),
            ));
            continue;
        }
        for (target, value) in targets.into_iter().zip(values) {
            body.push(fortran_data_assignment(target, value));
        }
    }
    Ok(Statement::new(StmtKind::Block(body)))
}

fn fortran_data_assignment(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}

/// An implied-do whose trip count is not known at walk time, kept as a loop.
struct FortranDataLoop {
    index_name: String,
    lower: Expression,
    upper: Expression,
    step: i64,
    /// Targets still written in terms of `index_name`.
    targets: Vec<Expression>,
}

/// One `data_var`: a bare name, a subscripted element, or an implied-do that
/// stands for several elements. `Some` when the implied-do has to stay a loop.
fn collect_fortran_data_targets(
    var: Pair<Rule>,
    out: &mut Vec<Expression>,
) -> Result<Option<FortranDataLoop>, String> {
    let inner: Vec<Pair<Rule>> = var.into_inner().filter(|p| meaningful(p)).collect();
    if let Some(implied) = inner
        .iter()
        .find(|p| p.as_rule() == Rule::data_implied_do)
    {
        return expand_fortran_data_implied_do(implied.clone(), out);
    }
    out.push(build_fortran_data_target(&inner)?);
    Ok(None)
}

/// `data (out(i), i = 1, n) /10, 20, 30/` — the values go into a temporary and
/// a counter walks them as the loop runs, because neither the trip count nor
/// which value lands where is known before execution.
fn build_fortran_data_loop(form: FortranDataLoop, values: Vec<Expression>) -> Statement {
    let cursor = format!("__fortran_data_cursor_{}", form.index_name);
    let store = format!("__fortran_data_values_{}", form.index_name);
    let value_count = values.len() as i64;
    let mut body = Vec::new();
    for target in form.targets {
        // The read is written the way FORTRAN writes one — a `Call` on a
        // 1-based cursor — because the store is a declared array and every
        // subscript pass will treat it as one. Reaching in with a raw 0-based
        // `Index` got the 1-based adjustment applied on top and read one
        // element early.
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&store)),
                args: vec![Argument::positional(Expression::ident(&cursor))],
                optional: false,
            }),
            by_ref: false,
        }));
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&cursor)],
            value: Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&cursor)),
                right: Box::new(Expression::int(1)),
            }),
            by_ref: false,
        }));
    }
    let cond = Expression::new(ExprKind::Binary {
        op: if form.step > 0 {
            BinOp::LtEq
        } else {
            BinOp::GtEq
        },
        left: Box::new(Expression::ident(&form.index_name)),
        right: Box::new(form.upper),
    });
    Statement::new(StmtKind::Block(vec![
        fortran_data_array(&store, fortran_array_expr(values), value_count),
        fortran_data_local(&cursor, Expression::int(1)),
        Statement::new(StmtKind::For {
            init: Some(Box::new(fortran_data_local(&form.index_name, form.lower))),
            cond: Some(cond),
            update: Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident(&form.index_name)),
                value: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(Expression::ident(&form.index_name)),
                    right: Box::new(Expression::int(form.step)),
                })),
            })),
            body,
        }),
    ]))
}

fn fortran_data_local(name: &str, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

/// The value store, declared with bounds so the subscript passes see a Fortran
/// array rather than a bare list.
fn fortran_data_array(name: &str, init: Expression, len: i64) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: Some(vec![Expression::int(len)]),
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

/// `name` → `Ident`, `name(args)` → `Call`.
fn build_fortran_data_target(parts: &[Pair<Rule>]) -> Result<Expression, String> {
    let name = parts
        .iter()
        .find(|p| p.as_rule() == Rule::identifier)
        .ok_or("missing name in data target")?
        .as_str();
    let mut args = Vec::new();
    for list in parts
        .iter()
        .filter(|p| p.as_rule() == Rule::argument_list)
    {
        for a in list.clone().into_inner() {
            if a.as_rule() == Rule::argument {
                let (name, value) = walk_argument_expr(a)?;
                args.push(Argument {
                    name,
                    value,
                    by_ref: false,
                    spread: false,
                });
            }
        }
    }
    if args.is_empty() {
        return Ok(Expression::ident(name));
    }
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args,
        optional: false,
    }))
}

/// `(a(i), i = lo, hi [, step])` — one target per trip.
///
/// The bounds must be known here, because the targets are separate assignment
/// statements rather than a loop. That matches the corpus (`i = 1, 5, 2`), and
/// the fold below covers the `2 + 0` spelling the suite also uses.
fn expand_fortran_data_implied_do(
    implied: Pair<Rule>,
    out: &mut Vec<Expression>,
) -> Result<Option<FortranDataLoop>, String> {
    let mut vars = Vec::new();
    let mut index_name = None;
    let mut bounds = Vec::new();
    for part in implied.into_inner().filter(|p| meaningful(p)) {
        match part.as_rule() {
            Rule::data_var_simple => {
                vars.push(part.into_inner().filter(|p| meaningful(p)).collect::<Vec<_>>())
            }
            Rule::identifier if index_name.is_none() => {
                index_name = Some(part.as_str().to_string())
            }
            Rule::expression => bounds.push(walk_expr(part)?),
            _ => {}
        }
    }
    let index_name = index_name.ok_or("missing index in data implied do")?;
    let step = match bounds.get(2) {
        Some(step) => {
            fortran_data_const_int(step).ok_or("non-constant implied-do step in data statement")?
        }
        None => 1,
    };
    if step == 0 {
        return Err("zero implied-do step in data statement".to_string());
    }
    let raw_lower = bounds.first().ok_or("missing implied-do lower bound")?;
    let raw_upper = bounds.get(1).ok_or("missing implied-do upper bound")?;
    let (Some(lower), Some(upper)) = (
        fortran_data_const_int(raw_lower),
        fortran_data_const_int(raw_upper),
    ) else {
        // `data (out(i), i = 1, n)` with `n` a named constant: the trip count is
        // not known here, so the targets cannot be separate statements. The
        // whole set becomes a loop instead, walking the value list with a
        // counter.
        let mut targets = Vec::with_capacity(vars.len());
        for parts in &vars {
            targets.push(build_fortran_data_target(parts)?);
        }
        return Ok(Some(FortranDataLoop {
            index_name,
            lower: raw_lower.clone(),
            upper: raw_upper.clone(),
            step,
            targets,
        }));
    };
    let mut index = lower;
    while (step > 0 && index <= upper) || (step < 0 && index >= upper) {
        for parts in &vars {
            let target = build_fortran_data_target(parts)?;
            out.push(substitute_fortran_ident_expr(
                &target,
                &index_name,
                &Expression::int(index),
            ));
        }
        index += step;
    }
    Ok(None)
}

/// A DATA bound, folded to a number. Literals plus the one arithmetic level the
/// suite spells (`2 + 0`); anything else is reported rather than guessed.
fn fortran_data_const_int(expr: &Expression) -> Option<i64> {
    if let Some(value) = fortran_literal_int(expr) {
        return Some(value);
    }
    // `i = 5, 1, -1` — a negative step is a UNARY MINUS over a literal, never a
    // negative literal, so folding only `Lit` reads every descending implied-do
    // as non-constant.
    if let ExprKind::Unary { op, expr: inner } = &expr.kind {
        let value = fortran_data_const_int(inner)?;
        return match op {
            UnaryOp::Neg => Some(-value),
            UnaryOp::Pos => Some(value),
            _ => None,
        };
    }
    let ExprKind::Binary { op, left, right } = &expr.kind else {
        return None;
    };
    let left = fortran_data_const_int(left)?;
    let right = fortran_data_const_int(right)?;
    match op {
        BinOp::Add => Some(left + right),
        BinOp::Sub => Some(left - right),
        BinOp::Mul => Some(left * right),
        BinOp::Div if right != 0 => Some(left / right),
        _ => None,
    }
}

/// One `data_value`, which is `expr` or the repeat form `count * expr`.
fn collect_fortran_data_values(
    value: Pair<Rule>,
    out: &mut Vec<Expression>,
) -> Result<(), String> {
    let parts: Vec<Pair<Rule>> = value.into_inner().filter(|p| meaningful(p)).collect();
    let first = walk_expr(parts.first().cloned().ok_or("empty data value")?)?;
    let Some(second) = parts.get(1) else {
        out.push(first);
        return Ok(());
    };
    let repeat = fortran_data_const_int(&first)
        .ok_or("non-constant repeat count in data statement")?
        .max(0);
    let value = walk_expr(second.clone())?;
    for _ in 0..repeat {
        out.push(value.clone());
    }
    Ok(())
}

/// Fortran 95 `FORALL (i = lo:hi[:step], …[, mask]) …` — indexed elementwise
/// assignment.
///
/// The grammar has carried the construct since it was added, and its own
/// comment says "walker treats it as a nested DO loop family" — but no walker
/// arm existed, so `_ => Ok(None)` dropped it and every FORALL body simply
/// never ran.
///
/// Each body statement gets its OWN full nest, because FORALL completes a
/// statement for every index before the next statement begins — running them
/// inside one nest would let a later assignment read what an earlier one wrote
/// for a different index. The mask, when present, guards the innermost body.
///
/// ⛔ Not buffered WITHIN one statement: `forall (i=1:n) a(i) = a(i+1)` reads
/// values this lowering may already have overwritten. Fortran evaluates every
/// right-hand side first; that needs a temporary per statement and is not done
/// here.
fn walk_forall(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut triplets = Vec::new();
    let mut mask = None;
    let mut body = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::forall_triplet => triplets.push(p),
            Rule::statement_line => body.extend(walk_statement_line_stmts(p)?),
            _ if is_expr_rule(p.as_rule()) && mask.is_none() => mask = Some(walk_expr(p)?),
            _ => {
                if let Some(statement) = walk_stmt(p)? {
                    body.push(statement);
                }
            }
        }
    }
    if triplets.is_empty() {
        return Ok(Statement::new(StmtKind::Block(body)));
    }

    let mut headers = Vec::with_capacity(triplets.len());
    for triplet in triplets {
        let parts: Vec<Pair<Rule>> = triplet.into_inner().filter(|p| meaningful(p)).collect();
        let name = parts.first().ok_or("missing forall index")?.as_str().to_string();
        let lower = walk_expr(parts.get(1).ok_or("missing forall lower bound")?.clone())?;
        let upper = walk_expr(parts.get(2).ok_or("missing forall upper bound")?.clone())?;
        let step = match parts.get(3) {
            Some(step) => walk_expr(step.clone())?,
            None => Expression::int(1),
        };
        headers.push((name, lower, upper, step));
    }

    let mut nests = Vec::with_capacity(body.len());
    for statement in body {
        let mut inner = vec![statement];
        if let Some(mask) = mask.clone() {
            inner = vec![Statement::new(StmtKind::If {
                cond: mask,
                then_body: inner,
                elifs: Vec::new(),
                else_body: None,
            })];
        }
        for (name, lower, upper, step) in headers.iter().rev() {
            inner = vec![build_fortran_counted_loop(
                name,
                lower.clone(),
                upper.clone(),
                step.clone(),
                inner,
            )];
        }
        nests.extend(inner);
    }
    Ok(Statement::new(StmtKind::Block(nests)))
}

/// F77 `do 100 i = 1, 4 … 100 continue` — a DO whose terminator is a LABEL.
///
/// The grammar deliberately gives this rule no terminator: pest cannot check
/// that the closing label matches, so the body and the `100 continue` are
/// absorbed as SIBLINGS of the header by the enclosing `line*`. A walker arm
/// never existed, so the header was dropped and the "body" simply ran once.
///
/// The loop is built here with an EMPTY body, wrapped in a marker naming the
/// terminator; `lower_fortran_labeled_do` then moves the siblings in. It can
/// do that because statement labels now survive as `Labeled` wrappers — the
/// same capture the GOTO work added.
fn walk_labeled_do(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut terminator = None;
    let mut var = None;
    let mut bounds = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::statement_label if terminator.is_none() => {
                terminator = Some(p.as_str().to_string())
            }
            Rule::identifier if var.is_none() => var = Some(p.as_str().to_string()),
            _ if is_expr_rule(p.as_rule()) => bounds.push(walk_expr(p)?),
            _ => {}
        }
    }
    let terminator = terminator.ok_or("missing labeled do terminator")?;
    let var = var.ok_or("missing labeled do index")?;
    let lower = bounds.first().cloned().ok_or("missing labeled do lower bound")?;
    let upper = bounds.get(1).cloned().ok_or("missing labeled do upper bound")?;
    let step = bounds.get(2).cloned().unwrap_or_else(|| Expression::int(1));
    Ok(Statement::new(StmtKind::Labeled {
        label: format!("{FORTRAN_LABELED_DO_PREFIX}{terminator}"),
        body: Box::new(build_fortran_counted_loop(&var, lower, upper, step, Vec::new())),
    }))
}

/// Marks a loop still waiting for the statements up to its terminator.
const FORTRAN_LABELED_DO_PREFIX: &str = "__fortran_labeled_do_";

/// Move each labelled DO's siblings into its body, up to its terminator.
fn lower_fortran_labeled_do(body: &mut Vec<Statement>) {
    for statement in body.iter_mut() {
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut lower_fortran_labeled_do);
    }
    let mut index = 0;
    while index < body.len() {
        let StmtKind::Labeled { label, .. } = &body[index].kind else {
            index += 1;
            continue;
        };
        let Some(terminator) = label.strip_prefix(FORTRAN_LABELED_DO_PREFIX) else {
            index += 1;
            continue;
        };
        let wanted = format!("{FORTRAN_LABEL_PREFIX}{terminator}");
        // Everything up to and INCLUDING the terminator line is the body; the
        // terminator is usually `continue`, which contributes nothing.
        let end = body[index + 1..]
            .iter()
            .position(|candidate| {
                matches!(&candidate.kind, StmtKind::Labeled { label, .. } if *label == wanted)
            })
            .map(|offset| index + 1 + offset);
        let Some(end) = end else {
            index += 1;
            continue;
        };
        let collected: Vec<Statement> = body.drain(index + 1..=end).collect();
        let StmtKind::Labeled { body: loop_stmt, .. } = &mut body[index].kind else {
            index += 1;
            continue;
        };
        if let StmtKind::For { body: loop_body, .. } = &mut loop_stmt.kind {
            *loop_body = collected;
            // A nested labelled DO was a SIBLING until this moment — the inner
            // `do 100` and its `100 continue` were both swept into the outer
            // body. Recursing beforehand cannot see them, so the inner loop is
            // resolved now that it finally lives somewhere.
            lower_fortran_labeled_do(loop_body);
        }
        // The wrapper has served its purpose; the loop stands on its own.
        body[index] = (**loop_stmt).clone();
        index += 1;
    }
}

/// `do name = lower, upper, step` as a counted `For`.
fn build_fortran_counted_loop(
    name: &str,
    lower: Expression,
    upper: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> Statement {
    // ⛔ `-1` is a NEGATION of the literal `1`, not a negative literal, so
    // asking `fortran_literal_int` reads every countdown as ascending and the
    // loop runs zero times. `fortran_step_is_negative` is the helper that
    // already knows this — the same trap the DATA implied-do step hit.
    let ascending = !fortran_step_is_negative(&step).unwrap_or(false);
    Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(name)],
            value: lower,
            by_ref: false,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: if ascending { BinOp::LtEq } else { BinOp::GtEq },
            left: Box::new(Expression::ident(name)),
            right: Box::new(upper),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(name)),
                right: Box::new(step),
            })),
        })),
        body,
    })
}

fn walk_do_concurrent(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut header = None;
    let mut body_parts = Vec::new();
    let mut body = Vec::new();

    for p in parts {
        match p.as_rule() {
            Rule::concurrent_header => header = Some(p),
            Rule::statement_line => body_parts.push(p),
            Rule::inline_statement_list => body.extend(walk_inline_statement_list(p)?),
            _ => {}
        }
    }

    body.extend(walk_body(body_parts.into_iter())?);

    let header = header.ok_or("missing concurrent header")?;
    let mut indices = Vec::new();
    let mut mask = None;

    for p in header.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::concurrent_index => {
                let mut inner = p.into_inner().filter(|item| meaningful(item));
                let var = inner
                    .next()
                    .ok_or("missing concurrent index variable")?
                    .as_str()
                    .to_string();
                let lower = walk_expr(inner.next().ok_or("missing concurrent lower bound")?)?;
                let upper = walk_expr(inner.next().ok_or("missing concurrent upper bound")?)?;
                let step = inner.next().map(walk_expr).transpose()?;
                indices.push((var, lower, upper, step));
            }
            _ if is_expr_rule(p.as_rule()) => {
                if mask.is_none() {
                    mask = Some(walk_expr(p)?);
                }
            }
            _ => {}
        }
    }

    if indices.is_empty() {
        return Ok(Statement::new(StmtKind::Block(body)));
    }

    let mut loop_body = body;
    if let Some(cond) = mask {
        loop_body = vec![Statement::new(StmtKind::If {
            cond,
            then_body: loop_body,
            elifs: Vec::new(),
            else_body: None,
        })];
    }

    for (var, start, end_e, step_expr) in indices.into_iter().rev() {
        let init = Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(var.clone()))],
            value: start,
            by_ref: false,
        })));
        let cond = Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
            op: BinOp::LtEq,
            right: Box::new(end_e),
        }));
        let step_value =
            step_expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
        let update = Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
            value: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::new(ExprKind::Ident(var))),
                op: BinOp::Add,
                right: Box::new(step_value),
            })),
        }));
        loop_body = vec![Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body: loop_body,
        })];
    }

    Ok(loop_body
        .into_iter()
        .next()
        .unwrap_or_else(|| Statement::new(StmtKind::Block(Vec::new()))))
}

fn walk_do_while(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut cond = None;
    let mut body_parts = Vec::new();
    let mut body = Vec::new();
    let mut label = None;
    for p in parts {
        if p.as_rule() == Rule::loop_label {
            // `watch: do while (…)` names the construct just like a counted
            // DO does, and `exit watch` has to be able to reach it.
            label = Some(p.as_str().to_string());
        } else if is_expr_rule(p.as_rule()) && cond.is_none() {
            cond = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::statement_line {
            body_parts.push(p);
        } else if p.as_rule() == Rule::inline_statement_list {
            body.extend(walk_inline_statement_list(p)?);
        }
        // skip kw_do, kw_while, kw_end
    }
    // If condition not found, emit "false" so the loop immediately exits (never infinite)
    let cond = cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(false))));
    body.extend(walk_body(body_parts.into_iter())?);
    Ok(label_fortran_loop(
        label,
        Statement::new(StmtKind::While {
            cond,
            body,
            else_body: None,
        }),
    ))
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
    let mut default_body: Option<Vec<Statement>> = None;
    for p in case_pairs {
        let mut conds = Vec::new();
        let mut cbody = Vec::new();
        let mut is_default = false;
        for c in p.into_inner().filter(|p| meaningful(p)) {
            match c.as_rule() {
                Rule::case_value_list => {
                    for cv in c.into_inner() {
                        if cv.as_rule() != Rule::case_value {
                            continue;
                        }
                        let cv_text = cv.as_str();
                        let cv_start = cv.as_span().start();
                        let cv_children: Vec<Pair<Rule>> =
                            cv.clone().into_inner().filter(|p| meaningful(p)).collect();
                        // Range: expr? ":" expr?  — two expressions separated by ":"
                        // cv children for `e1:e2` are two expr pairs; for `e1` just one.
                        let cv_exprs: Vec<Pair<Rule>> = cv_children
                            .into_iter()
                            .filter(|p| {
                                is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression
                            })
                            .collect();
                        if cv_exprs.len() >= 2 {
                            let from = walk_expr(cv_exprs[0].clone())?;
                            let to = walk_expr(cv_exprs[1].clone())?;
                            conds.push(CaseCondition::Range { from, to });
                        } else if let Some(first) = cv_exprs.into_iter().next() {
                            // One expression can still be a RANGE: `case (:hi)`
                            // and `case (lo:)` are open-ended, and an open end
                            // is a one-sided comparison against the bound that
                            // IS written. The ":" is a bare literal in the
                            // grammar, so it has no child pair — locate it by
                            // position, and only OUTSIDE the expression's own
                            // span, or `case (':')` would read as a range.
                            let expr_start = first.as_span().start() - cv_start;
                            let expr_end = first.as_span().end() - cv_start;
                            let colon = cv_text
                                .char_indices()
                                .find(|(i, ch)| {
                                    *ch == ':' && (*i < expr_start || *i >= expr_end)
                                })
                                .map(|(i, _)| i);
                            let value = walk_expr(first)?;
                            conds.push(match colon {
                                // `case (:hi)` — every value up to and including hi.
                                Some(at) if at < expr_start => CaseCondition::Comparison {
                                    op: ComparisonOp::LtEq,
                                    expr: value,
                                },
                                // `case (lo:)` — lo and everything above it.
                                Some(_) => CaseCondition::Comparison {
                                    op: ComparisonOp::GtEq,
                                    expr: value,
                                },
                                None => CaseCondition::Value(value),
                            });
                        }
                    }
                }
                Rule::kw_default => {
                    is_default = true;
                }
                Rule::statement_line => {
                    for s in c.into_inner().filter(|p| meaningful(p)) {
                        if let Some(st) = walk_stmt(s)? {
                            cbody.push(st);
                        }
                    }
                }
                _ => {
                    if c.as_str().trim().to_ascii_lowercase() == "default" {
                        is_default = true;
                    }
                }
            }
        }
        if is_default {
            default_body = Some(cbody);
        } else {
            cases.push(SwitchCase {
                conditions: conds,
                body: cbody,
            });
        }
    }
    Ok(Statement::new(StmtKind::Switch {
        expr,
        cases,
        default: default_body,
    }))
}

fn walk_select_type(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut selector: Option<Expression> = None;
    let mut clauses = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        if selector.is_none() && (is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression) {
            selector = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::select_type_clause {
            clauses.push(p);
        }
    }
    let selector = selector.unwrap_or_else(|| Expression::null());
    let mut branches: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;
    for clause in clauses {
        let header = fortran_clause_header(&clause);
        let body = walk_fortran_clause_body(clause)?;
        if header.contains("class default") {
            default_body = Some(body);
            continue;
        }
        let type_name = fortran_clause_paren_text(&header)
            .map(fortran_canonical_select_type_name)
            .unwrap_or_else(|| "object".to_string());
        // `type is` matches the EXACT dynamic type; `class is` matches that
        // type or any extension of it. Both used to lower to the same test, so
        // `type is (Base)` claimed a `Child`. The distinction is carried on the
        // marker and resolved once the type hierarchy is visible.
        let exact = header.contains("type is");
        branches.push((
            build_fortran_select_type_condition(selector.clone(), &type_name, exact),
            body,
        ));
    }
    // The chain is marked so the specificity pass can find it: Fortran selects
    // the MOST SPECIFIC matching branch, not the first one written.
    Ok(Statement::new(StmtKind::Labeled {
        label: FORTRAN_SELECT_TYPE_MARKER.to_string(),
        body: Box::new(fortran_if_from_branches(branches, default_body)),
    }))
}

fn walk_select_rank(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut selector: Option<Expression> = None;
    let mut clauses = Vec::new();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        if selector.is_none() && (is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression) {
            selector = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::select_rank_clause {
            clauses.push(p);
        }
    }
    let selector = selector.unwrap_or_else(|| Expression::null());
    let rank_expr = build_fortran_rank_expr(selector);
    let mut branches: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;
    for clause in clauses {
        let header = fortran_clause_header(&clause);
        let body = walk_fortran_clause_body(clause)?;
        if header.contains("rank default") {
            default_body = Some(body);
            continue;
        }
        let Some(rank_text) = fortran_clause_paren_text(&header) else {
            continue;
        };
        if rank_text.trim() == "*" {
            default_body = Some(body);
            continue;
        }
        let rank_value = parse_fortran_expression_text(rank_text.trim())
            .unwrap_or_else(|_| Expression::int(0));
        branches.push((
            Expression::new(ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(rank_expr.clone()),
                right: Box::new(rank_value),
            }),
            body,
        ));
    }
    Ok(fortran_if_from_branches(branches, default_body))
}

fn fortran_clause_header(pair: &Pair<Rule>) -> String {
    pair.as_str()
        .lines()
        .next()
        .unwrap_or("")
        .trim()
        .to_ascii_lowercase()
}

fn fortran_clause_paren_text(header: &str) -> Option<&str> {
    let start = header.find('(')?;
    let end = header.rfind(')')?;
    (end > start).then_some(header[start + 1..end].trim())
}

fn walk_fortran_clause_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for child in pair.into_inner().filter(|p| meaningful(p)) {
        if child.as_rule() != Rule::statement_line {
            continue;
        }
        for stmt_pair in child.into_inner().filter(|p| meaningful(p)) {
            if let Some(stmt) = walk_stmt(stmt_pair)? {
                body.push(stmt);
            }
        }
    }
    let mut type_env = HashMap::new();
    lower_fortran_body_intrinsics_with_env(&mut body, &mut type_env);
    Ok(body)
}

fn fortran_canonical_select_type_name(type_name: &str) -> String {
    let lower = type_name.trim().to_ascii_lowercase();
    if lower.starts_with("integer") {
        "integer".to_string()
    } else if lower.starts_with("real")
        || lower.starts_with("double precision")
        || lower.starts_with("double")
    {
        "number".to_string()
    } else if lower.starts_with("logical") {
        "boolean".to_string()
    } else if lower.starts_with("character") {
        "string".to_string()
    } else if lower.starts_with("complex") {
        "object".to_string()
    } else if let Some(inner) = lower.strip_prefix("type(").and_then(|s| s.strip_suffix(')')) {
        inner.to_string()
    } else if let Some(inner) = lower.strip_prefix("class(").and_then(|s| s.strip_suffix(')')) {
        inner.to_string()
    } else {
        type_name.trim().to_string()
    }
}

/// Marks a `select type` chain still awaiting the specificity rules.
const FORTRAN_SELECT_TYPE_MARKER: &str = "__fortran_select_type";
/// Prefix on a type name that must match EXACTLY (`type is`).
const FORTRAN_EXACT_TYPE_PREFIX: &str = "=";

fn build_fortran_select_type_condition(
    selector: Expression,
    type_name: &str,
    exact: bool,
) -> Expression {
    let type_name = if exact {
        format!("{FORTRAN_EXACT_TYPE_PREFIX}{type_name}")
    } else {
        type_name.to_string()
    };
    Expression::new(ExprKind::IsType {
        expr: Box::new(selector),
        type_name,
    })
}

/// Apply Fortran's `select type` selection rules to every marked chain.
///
/// Two rules, both invisible until typed allocation started working (with every
/// object stuck as its declared type, no branch after the first could ever be
/// reached):
///
/// 1. `type is (T)` matches the EXACT dynamic type — a `Child` is not a `Base`.
///    Expressed as "is a T, and is none of T's extensions", which is what the
///    hierarchy makes answerable.
/// 2. The MOST SPECIFIC matching branch wins, not the first written. Rather
///    than reorder, each `class is (T)` also excludes the candidates in its own
///    chain that extend it — mutually exclusive conditions give the right
///    answer in any order.
fn lower_fortran_select_type_specificity(body: &mut Vec<Statement>) {
    let mut parents = HashMap::new();
    collect_fortran_type_parents(body, &mut parents);
    rewrite_fortran_select_type_chains(body, &parents, None);
    strip_fortran_alloc_type_markers(body);
}

/// Marks the intrinsic dynamic type a typed `allocate` gave a variable.
/// `__fortran_alloc_type:val:integer`.
const FORTRAN_ALLOC_TYPE_MARKER: &str = "__fortran_alloc_type:";

/// The intrinsic dynamic type of every unlimited polymorphic that a typed
/// `allocate` named — the ONLY record of it, since an `integer` and a `real`
/// are the same f64 once the program is running.
///
/// A variable allocated twice with DIFFERENT types is dropped rather than
/// guessed at: its dynamic type genuinely depends on which allocation ran, and
/// answering from one of them would be a coin flip dressed as a fact.
///
/// ⛔ ONE SCOPE. It walks the nested bodies of `if`/`do` — an `allocate` under a
/// branch is the same scope — but stops at a procedure, because a `val`
/// allocated in one subroutine says nothing about a `val` in the next, and a
/// tree-wide map would let one fold a `select type` it never allocated.
fn collect_fortran_alloc_types(body: &[Statement], out: &mut HashMap<String, Option<String>>) {
    for statement in body {
        if let StmtKind::Labeled { label, .. } = &statement.kind {
            if let Some(rest) = label.strip_prefix(FORTRAN_ALLOC_TYPE_MARKER) {
                if let Some((var, type_name)) = rest.split_once(':') {
                    out.entry(var.to_string())
                        .and_modify(|seen| {
                            if seen.as_deref() != Some(type_name) {
                                *seen = None;
                            }
                        })
                        .or_insert_with(|| Some(type_name.to_string()));
                }
            }
        }
        if matches!(
            statement.kind,
            StmtKind::FunctionDecl { .. }
                | StmtKind::ModuleDecl { .. }
                | StmtKind::ClassDecl { .. }
                | StmtKind::StructDecl { .. }
        ) {
            continue;
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_alloc_types(stmts, out)
        });
    }
}

/// The markers are spent once the chains have read them. They become an empty
/// block rather than being removed, so this can run over the `&mut [Statement]`
/// slices the nested-body traversal hands out.
fn strip_fortran_alloc_type_markers(body: &mut [Statement]) {
    for statement in body.iter_mut() {
        strip_fortran_alloc_type_markers_in_statement(statement);
    }
}

fn strip_fortran_alloc_type_markers_in_statement(statement: &mut Statement) {
    if let StmtKind::Labeled { label, .. } = &statement.kind {
        if label.starts_with(FORTRAN_ALLOC_TYPE_MARKER) {
            *statement = Statement::new(StmtKind::Block(Vec::new()));
            return;
        }
    }
    match &mut statement.kind {
        StmtKind::FunctionDecl { body, .. } => strip_fortran_alloc_type_markers(body),
        StmtKind::ModuleDecl { members, .. }
        | StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. } => {
            for member in members.iter_mut() {
                if let ClassMember::Method(inner) | ClassMember::NestedType(inner) = member {
                    strip_fortran_alloc_type_markers_in_statement(inner);
                }
            }
        }
        _ => {}
    }
    for_each_fortran_nested_body_mut(&mut statement.kind, &mut |stmts| {
        strip_fortran_alloc_type_markers(stmts)
    });
}

/// Resolve a `select type` chain whose selector has a KNOWN intrinsic dynamic
/// type: keep the branch that names it and drop the rest.
///
/// This is not an optimisation — it is the only correct answer available. The
/// runtime test behind `type is (integer)` asks whether the value is a number,
/// and a `real` is one too, so every intrinsic branch claims every intrinsic
/// value and the first one written always wins.
fn fold_fortran_select_type_by_alloc(
    chain: &Statement,
    alloc_types: &HashMap<String, Option<String>>,
) -> Option<Statement> {
    let StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    } = &chain.kind
    else {
        return None;
    };
    let selector_type = |expr: &Expression| -> Option<String> {
        let ExprKind::IsType { expr: selector, .. } = &expr.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &selector.kind else {
            return None;
        };
        alloc_types.get(&name.to_ascii_lowercase())?.clone()
    };
    let dynamic = selector_type(cond)?;
    let branch_matches = |expr: &Expression| -> bool {
        let ExprKind::IsType { type_name, .. } = &expr.kind else {
            return false;
        };
        type_name
            .strip_prefix(FORTRAN_EXACT_TYPE_PREFIX)
            .unwrap_or(type_name)
            .eq_ignore_ascii_case(&dynamic)
    };
    if branch_matches(cond) {
        return Some(Statement::new(StmtKind::Block(then_body.clone())));
    }
    for (elif_cond, elif_body) in elifs {
        // ⛔ Every arm has to be one of ours. A chain that mixes in a condition
        // this pass cannot read is not a chain it may collapse.
        selector_type(elif_cond)?;
        if branch_matches(elif_cond) {
            return Some(Statement::new(StmtKind::Block(elif_body.clone())));
        }
    }
    Some(Statement::new(StmtKind::Block(
        else_body.clone().unwrap_or_default(),
    )))
}


fn collect_fortran_type_parents(body: &[Statement], out: &mut HashMap<String, Vec<String>>) {
    for statement in body {
        if let StmtKind::ClassDecl { name, parents, .. } = &statement.kind {
            out.insert(
                name.to_ascii_lowercase(),
                parents.iter().map(|p| p.to_ascii_lowercase()).collect(),
            );
        }
        if let StmtKind::FunctionDecl { body, .. } = &statement.kind {
            collect_fortran_type_parents(body, out);
        }
        if let StmtKind::ModuleDecl { members, .. } | StmtKind::ClassDecl { members, .. } =
            &statement.kind
        {
            for member in members {
                if let ClassMember::NestedType(inner) | ClassMember::Method(inner) = member {
                    collect_fortran_type_parents(std::slice::from_ref(inner), out);
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_type_parents(stmts, out)
        });
    }
}

/// Whether `descendant` extends `ancestor`, transitively.
fn fortran_type_extends(
    descendant: &str,
    ancestor: &str,
    parents: &HashMap<String, Vec<String>>,
) -> bool {
    let mut queue: Vec<String> = parents.get(descendant).cloned().unwrap_or_default();
    let mut seen = HashSet::new();
    while let Some(next) = queue.pop() {
        if next == ancestor {
            return true;
        }
        if !seen.insert(next.clone()) {
            continue;
        }
        queue.extend(parents.get(&next).cloned().unwrap_or_default());
    }
    false
}

fn rewrite_fortran_select_type_chains(
    body: &mut [Statement],
    parents: &HashMap<String, Vec<String>>,
    scope_alloc_types: Option<&HashMap<String, Option<String>>>,
) {
    // The enclosing scope's typed allocations, or this body's own when it IS
    // the scope. A procedure starts a fresh one — see
    // `collect_fortran_alloc_types`.
    let owned_alloc_types = scope_alloc_types.is_none().then(|| {
        let mut collected = HashMap::new();
        collect_fortran_alloc_types(body, &mut collected);
        collected
    });
    let alloc_types = scope_alloc_types.unwrap_or_else(|| {
        owned_alloc_types
            .as_ref()
            .expect("collected when no scope was inherited")
    });
    for statement in body.iter_mut() {
        // ⛔ Procedure and class bodies too: the general traversal skips them
        // (a GOTO cannot leave its procedure), but a `select type` inside one
        // would keep its marker, and with it the `=` prefix that no type name
        // can match — the chain would then fall through every branch.
        match &mut statement.kind {
            StmtKind::FunctionDecl { body, .. } => {
                rewrite_fortran_select_type_chains(body, parents, None)
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                for member in members.iter_mut() {
                    if let ClassMember::Method(inner) | ClassMember::NestedType(inner) = member {
                        rewrite_fortran_select_type_chains(
                            std::slice::from_mut(inner),
                            parents,
                            None,
                        );
                    }
                }
            }
            _ => {}
        }
        for_each_fortran_nested_body_mut(&mut statement.kind, &mut |stmts| {
            rewrite_fortran_select_type_chains(stmts, parents, Some(alloc_types))
        });
        let StmtKind::Labeled { label, body: inner } = &mut statement.kind else {
            continue;
        };
        if label != FORTRAN_SELECT_TYPE_MARKER {
            continue;
        }
        let mut chain = (**inner).clone();
        // A known intrinsic dynamic type answers the chain outright; the
        // specificity rules are for the hierarchy questions that remain.
        if let Some(folded) = fold_fortran_select_type_by_alloc(&chain, alloc_types) {
            *statement = folded;
            continue;
        }
        apply_fortran_select_type_rules(&mut chain, parents);
        *statement = chain;
    }
}

fn apply_fortran_select_type_rules(
    chain: &mut Statement,
    parents: &HashMap<String, Vec<String>>,
) {
    let StmtKind::If {
        cond, elifs, ..
    } = &mut chain.kind
    else {
        return;
    };
    // Every type named in this chain — a `class is` only has to beat the
    // candidates it is actually competing with.
    let mut candidates: Vec<String> = Vec::new();
    let mut named = |expr: &Expression| {
        if let ExprKind::IsType { type_name, .. } = &expr.kind {
            candidates.push(
                type_name
                    .strip_prefix(FORTRAN_EXACT_TYPE_PREFIX)
                    .unwrap_or(type_name)
                    .to_ascii_lowercase(),
            );
        }
    };
    named(cond);
    for (elif_cond, _) in elifs.iter() {
        named(elif_cond);
    }

    let refine = |expr: &mut Expression| {
        let ExprKind::IsType { expr: selector, type_name } = &expr.kind else {
            return;
        };
        let exact = type_name.starts_with(FORTRAN_EXACT_TYPE_PREFIX);
        let bare = type_name
            .strip_prefix(FORTRAN_EXACT_TYPE_PREFIX)
            .unwrap_or(type_name)
            .to_string();
        let lower = bare.to_ascii_lowercase();
        // An EXACT match excludes every extension there is; a `class is`
        // excludes only the rival candidates that are more specific than it.
        let rivals: Vec<String> = if exact {
            parents
                .keys()
                .filter(|other| fortran_type_extends(other, &lower, parents))
                .cloned()
                .collect()
        } else {
            candidates
                .iter()
                .filter(|other| *other != &lower && fortran_type_extends(other, &lower, parents))
                .cloned()
                .collect()
        };
        let mut refined = Expression::new(ExprKind::IsType {
            expr: selector.clone(),
            type_name: bare,
        });
        for rival in rivals {
            refined = Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(refined),
                right: Box::new(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(Expression::new(ExprKind::IsType {
                        expr: selector.clone(),
                        type_name: rival,
                    })),
                })),
            });
        }
        *expr = refined;
    };
    refine(cond);
    for (elif_cond, _) in elifs.iter_mut() {
        refine(elif_cond);
    }
}

/// `RANK(x)` — and the selector of a `SELECT RANK`, which is the same question.
///
/// An ordinary call to the `rank` intrinsic, which the profile already binds:
/// `rank = { emit = "common:collections.rank" }`. Rank is a runtime property
/// here, not a static one — every `SELECT RANK` in the corpus dispatches on an
/// assumed-rank dummy (`integer, intent(in) :: x(..)`), so the same subroutine
/// is reached with a scalar and with an array and the walker cannot know which.
///
/// No AST node of its own: a rank is a fact about a collection VALUE, so it
/// belongs with the other collection primitives, where a language that has no
/// such surface never has to carry a variant it cannot produce.
fn build_fortran_rank_expr(selector: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("rank")),
        args: vec![Argument::positional(selector)],
        optional: false,
    })
}

fn fortran_if_from_branches(
    branches: Vec<(Expression, Vec<Statement>)>,
    default_body: Option<Vec<Statement>>,
) -> Statement {
    let mut iter = branches.into_iter();
    let Some((cond, then_body)) = iter.next() else {
        return Statement::new(StmtKind::Block(default_body.unwrap_or_default()));
    };
    Statement::new(StmtKind::If {
        cond,
        then_body,
        elifs: iter.collect(),
        else_body: default_body,
    })
}

fn walk_print(pair: Pair<Rule>) -> Result<Statement, String> {
    let raw = pair.as_str().trim_start().to_ascii_lowercase();
    let mut args = Vec::new();
    let mut explicit_format = false;
    let advance = true;
    let mut format_spec = None;

    match pair.as_rule() {
        Rule::print_statement => {
            let list_directed = raw
                .strip_prefix("print")
                .map(|rest| rest.trim_start().starts_with('*'))
                .unwrap_or(false);
            let mut skipped_format = false;
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if !list_directed && !skipped_format && p.as_rule() == Rule::string_literal {
                    explicit_format = true;
                    format_spec = Some(parse_fortran_string_literal_text(p.as_str()));
                    skipped_format = true;
                    continue;
                }
                if is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression {
                    args.push(walk_expr(p)?);
                }
            }
        }
        _ => {}
    }

    let text = build_fortran_text_expr(&args, explicit_format, format_spec.as_deref());
    let callee = if advance {
        "__fortran_emitln"
    } else {
        "__fortran_emit"
    };
    Ok(Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Call {
            callee: Box::new(Expression::ident(callee)),
            args: vec![Argument::positional(text)],
            optional: false,
        },
    ))))
}

fn walk_write(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut args = Vec::new();
    let mut explicit_format = false;
    let mut advance = true;
    let mut file_number = None;
    let mut format_spec = None;
    let mut namelist_group = None;
    let mut positional_spec_index = 0usize;

    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::write_spec => parse_fortran_write_spec(
                &p,
                &mut explicit_format,
                &mut advance,
                &mut file_number,
                &mut format_spec,
                &mut namelist_group,
                &mut positional_spec_index,
            )?,
            rule if is_expr_rule(rule) || rule == Rule::expression => args.push(walk_expr(p)?),
            _ => {}
        }
    }

    if let (Some(file_number), Some(group)) = (file_number.clone(), namelist_group) {
        return Ok(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__fortran_namelist_write")),
                args: vec![
                    Argument::positional(file_number),
                    Argument::positional(Expression::string(&group)),
                ],
                optional: false,
            },
        ))));
    }

    let text = build_fortran_text_expr(&args, explicit_format, format_spec.as_deref());
    if let Some(file_number) = file_number {
        if explicit_format {
            Ok(Statement::new(StmtKind::PrintFile {
                file_number,
                items: vec![text],
            }))
        } else {
            Ok(Statement::new(StmtKind::WriteFile {
                file_number,
                items: args,
            }))
        }
    } else {
        let callee = if advance {
            "__fortran_emitln"
        } else {
            "__fortran_emit"
        };
        Ok(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args: vec![Argument::positional(text)],
                optional: false,
            },
        ))))
    }
}

fn walk_read(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts = pair
        .into_inner()
        .filter(|p| meaningful(p))
        .collect::<Vec<_>>();
    let mut variables = Vec::new();
    let mut explicit_format = false;
    let mut advance = true;
    let mut file_number = None;
    let mut format_spec = None;
    let mut namelist_group = None;
    let mut iostat_target = None;
    let mut positional_spec_index = 0usize;

    for p in &parts {
        match p.as_rule() {
            Rule::write_spec => parse_fortran_write_spec(
                p,
                &mut explicit_format,
                &mut advance,
                &mut file_number,
                &mut format_spec,
                &mut namelist_group,
                &mut positional_spec_index,
            )?,
            rule if is_expr_rule(rule) || rule == Rule::expression => {
                variables.push(walk_expr(p.clone())?)
            }
            _ => {}
        }
    }

    for p in parts {
        if p.as_rule() != Rule::write_spec {
            continue;
        }
        let raw = p.as_str().trim().to_ascii_lowercase();
        if !raw.starts_with("iostat") {
            continue;
        }
        for item in p.into_inner().filter(|child| meaningful(child)) {
            if is_expr_rule(item.as_rule()) || item.as_rule() == Rule::expression {
                iostat_target = Some(walk_expr(item)?);
            }
        }
    }

    if let Some(group) = namelist_group {
        let mut args = vec![
            Argument::positional(file_number.unwrap_or_else(|| Expression::int(0))),
            Argument::positional(Expression::string(&group)),
        ];
        if let Some(iostat_target) = iostat_target {
            args.push(Argument::positional(iostat_target));
        }
        return Ok(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__fortran_namelist_read")),
                args,
                optional: false,
            },
        ))));
    }

    let read = Statement::new(StmtKind::InputFile {
        file_number: file_number.unwrap_or_else(|| Expression::int(0)),
        variables,
    });
    // `InputFile` carries only the unit and the targets, and whether the unit is
    // an INTERNAL file is a question about the unit's declared type — which is
    // known later, not here. So the format travels alongside as a marker, the
    // same protocol `__fortran_goto` uses, and the pass that does know the type
    // consumes it. It is ALWAYS consumed: `lower_fortran_internal_reads` strips
    // the marker whether or not it rewrites the read.
    let Some(format_spec) = format_spec.filter(|_| explicit_format) else {
        return Ok(read);
    };
    Ok(Statement::new(StmtKind::Block(vec![
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(FORTRAN_READ_FORMAT_MARKER)),
            args: vec![Argument::positional(Expression::string(&format_spec))],
            optional: false,
        }))),
        read,
    ])))
}

/// Carries a read's FORMAT from the walker to the pass that knows whether the
/// unit is an internal file.
const FORTRAN_READ_FORMAT_MARKER: &str = "__fortran_read_format";

fn parse_fortran_string_literal_text(raw: &str) -> String {
    // A KIND prefix (`c_char_"abc"`, `ascii_'x'`) sits before the opening quote.
    // The AST does not distinguish character kinds, so it is dropped — here, the
    // single place both the expression arm and the format-spec reader go
    // through, so neither can start slicing at the wrong byte.
    let Some(open) = raw.find(['\'', '"']) else {
        return raw.to_string();
    };
    if raw.len() < open + 2 {
        return raw.to_string();
    }
    raw[open + 1..raw.len() - 1]
        .replace("''", "'")
        .replace("\"\"", "\"")
}

fn parse_fortran_write_spec(
    spec: &Pair<Rule>,
    explicit_format: &mut bool,
    advance: &mut bool,
    file_number: &mut Option<Expression>,
    format_spec: &mut Option<String>,
    namelist_group: &mut Option<String>,
    positional_spec_index: &mut usize,
) -> Result<(), String> {
    let raw = spec.as_str().trim();
    if raw.eq("*") {
        *positional_spec_index += 1;
        return Ok(());
    }

    let lowered = raw.to_ascii_lowercase();
    if lowered.starts_with("advance") {
        if lowered.contains("'no'") || lowered.contains("\"no\"") {
            *advance = false;
        }
        return Ok(());
    }

    let mut named_spec = None;
    let mut expr = None;
    for item in spec.clone().into_inner().filter(|p| meaningful(p)) {
        match item.as_rule() {
            Rule::identifier => named_spec = Some(item.as_str().to_ascii_lowercase()),
            rule if is_expr_rule(rule) || rule == Rule::expression => expr = Some(walk_expr(item)?),
            _ => {}
        }
    }

    match named_spec.as_deref() {
        Some("unit") => {
            if let Some(expr) = expr {
                *file_number = Some(expr);
            }
        }
        Some("fmt") => {
            *explicit_format = true;
            if raw.contains('"') || raw.contains('\'') {
                if let Some(eq_idx) = raw.find('=') {
                    *format_spec =
                        Some(parse_fortran_string_literal_text(raw[eq_idx + 1..].trim()));
                }
            }
        }
        Some("nml") => {
            if let Some(Expression {
                kind: ExprKind::Ident(name),
                ..
            }) = expr
            {
                *namelist_group = Some(name);
            }
        }
        Some(_) => {}
        None => {
            if *positional_spec_index == 0 {
                *file_number = expr;
            } else if *positional_spec_index == 1 {
                *explicit_format = true;
                if raw.starts_with('"') || raw.starts_with('\'') {
                    *format_spec = Some(parse_fortran_string_literal_text(raw));
                }
            }
            *positional_spec_index += 1;
        }
    }

    Ok(())
}

#[derive(Debug, Clone)]
enum FortranFormatChunk {
    Data {
        descriptor: String,
        repeat: usize,
        width: Option<usize>,
        precision: Option<usize>,
    },
    Spaces(usize),
    Newline,
    Literal(String),
}

fn build_fortran_text_expr(
    args: &[Expression],
    explicit_format: bool,
    format_spec: Option<&str>,
) -> Expression {
    if explicit_format {
        if let Some(format_spec) = format_spec {
            if let Some(formatted) = build_fortran_formatted_io_text(args, format_spec) {
                return formatted;
            }
        }
    }
    build_fortran_io_text(args, explicit_format)
}

fn build_fortran_io_text(args: &[Expression], explicit_format: bool) -> Expression {
    let Some(first) = args.first() else {
        return Expression::string("");
    };
    let mut parts = vec![stringify_fortran_io_expr(first.clone())];

    for part in args.iter().skip(1).cloned().map(stringify_fortran_io_expr) {
        if explicit_format {
            parts.push(part);
        } else {
            parts.push(Expression::string(" "));
            parts.push(part);
        }
    }
    concat_fortran_io_parts(parts)
}

fn is_fortran_float_descriptor(descriptor: &str) -> bool {
    matches!(descriptor, "e" | "es" | "en" | "f" | "g" | "d")
}

fn build_fortran_formatted_io_text(args: &[Expression], format_spec: &str) -> Option<Expression> {
    let chunks = parse_fortran_format_chunks(format_spec)?;
    let mut arg_index = 0usize;
    let mut text_parts = Vec::new();
    let mut previous_chunk_was_data = false;

    for chunk in chunks {
        let (piece, chunk_is_data) = match chunk {
            FortranFormatChunk::Spaces(count) => (Expression::string(&" ".repeat(count)), false),
            FortranFormatChunk::Newline => (Expression::string("\n"), false),
            FortranFormatChunk::Literal(text) => (Expression::string(&text), false),
            FortranFormatChunk::Data {
                descriptor,
                repeat,
                width,
                precision,
            } => {
                // Float descriptors (e, es, f, g, d) include a sign-position space
                // when following another data item (Fortran field-width sign convention).
                if previous_chunk_was_data && is_fortran_float_descriptor(&descriptor) {
                    text_parts.push(Expression::string(" "));
                }
                let expand_single_array = repeat > 1 && args.len().saturating_sub(arg_index) == 1;
                if !expand_single_array && args.len().saturating_sub(arg_index) < repeat {
                    return None;
                }
                let segment = if expand_single_array {
                    build_fortran_formatted_array_segment(
                        args[arg_index].clone(),
                        &descriptor,
                        repeat,
                        width,
                        precision,
                    )?
                } else {
                    let mut segment: Option<Expression> = None;
                    for offset in 0..repeat {
                        let formatted = build_fortran_formatted_value_expr(
                            args[arg_index + offset].clone(),
                            &descriptor,
                            width,
                            precision,
                        )?;
                        segment = Some(match segment {
                            Some(existing) => concat_fortran_io_parts(vec![
                                existing,
                                Expression::string(" "),
                                formatted,
                            ]),
                            None => formatted,
                        });
                    }
                    segment.unwrap_or_else(|| Expression::string(""))
                };
                arg_index += if expand_single_array { 1 } else { repeat };
                (segment, true)
            }
        };
        text_parts.push(piece);
        previous_chunk_was_data = chunk_is_data;
    }

    if arg_index != args.len() {
        return None;
    }

    Some(concat_fortran_io_parts(text_parts))
}

fn build_fortran_formatted_array_segment(
    array_expr: Expression,
    descriptor: &str,
    repeat: usize,
    width: Option<usize>,
    precision: Option<usize>,
) -> Option<Expression> {
    let sliced = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_expr),
            field: "slice".to_string(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(Expression::int(0)),
            Argument::positional(Expression::int(repeat as i64)),
        ],
        optional: false,
    });
    let item_name = "__fortran_formatted_item";
    let mapped = build_fortran_array_map(
        sliced,
        build_fortran_formatted_value_expr(
            Expression::ident(item_name),
            descriptor,
            width,
            precision,
        )?,
        false,
        item_name,
        "__fortran_formatted_index",
    );

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(mapped),
            field: "join".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::string(" "))],
        optional: false,
    }))
}

fn build_fortran_formatted_value_expr(
    expr: Expression,
    descriptor: &str,
    _width: Option<usize>,
    precision: Option<usize>,
) -> Option<Expression> {
    let formatted = match descriptor {
        "a" | "l" => stringify_fortran_io_expr(expr),
        "i" => stringify_fortran_io_expr(expr),
        "f" => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("tofixed")),
            args: vec![
                Argument::positional(expr),
                Argument::positional(Expression::int(precision.unwrap_or(6) as i64)),
            ],
            optional: false,
        }),
        "e" | "es" | "d" | "g" => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("toexponential")),
            args: vec![
                Argument::positional(expr),
                Argument::positional(Expression::int(precision.unwrap_or(6) as i64)),
            ],
            optional: false,
        }),
        _ => return None,
    };

    Some(formatted)
}

fn parse_fortran_format_chunks(format_spec: &str) -> Option<Vec<FortranFormatChunk>> {
    let mut source = format_spec.trim();
    if source.starts_with('(') && source.ends_with(')') && source.len() >= 2 {
        source = &source[1..source.len() - 1];
    }

    let bytes = source.as_bytes();
    let mut index = 0usize;
    let mut chunks = Vec::new();

    while index < bytes.len() {
        let ch = bytes[index] as char;
        if ch.is_ascii_whitespace() || ch == ',' {
            index += 1;
            continue;
        }
        if ch == '/' {
            chunks.push(FortranFormatChunk::Newline);
            index += 1;
            continue;
        }
        if ch == '\'' || ch == '"' {
            let quote = ch;
            index += 1;
            let start = index;
            while index < bytes.len() && bytes[index] as char != quote {
                index += 1;
            }
            if index > bytes.len() {
                return None;
            }
            chunks.push(FortranFormatChunk::Literal(
                source[start..index].to_string(),
            ));
            if index < bytes.len() {
                index += 1;
            }
            continue;
        }

        let repeat_start = index;
        while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
            index += 1;
        }
        let repeat_prefix = if repeat_start < index {
            source[repeat_start..index].parse::<usize>().ok()
        } else {
            None
        };
        if index >= bytes.len() {
            break;
        }

        let mut descriptor = String::new();
        descriptor.push((bytes[index] as char).to_ascii_lowercase());
        index += 1;
        if descriptor == "e" && index < bytes.len() {
            let suffix = (bytes[index] as char).to_ascii_lowercase();
            if suffix == 's' {
                descriptor.push(suffix);
                index += 1;
            }
        }

        if descriptor == "x" {
            chunks.push(FortranFormatChunk::Spaces(repeat_prefix.unwrap_or(1)));
            continue;
        }

        let width_start = index;
        while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
            index += 1;
        }
        let width = if width_start < index {
            source[width_start..index].parse::<usize>().ok()
        } else {
            None
        };

        let precision = if index < bytes.len() && bytes[index] as char == '.' {
            index += 1;
            let precision_start = index;
            while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
                index += 1;
            }
            if precision_start < index {
                source[precision_start..index].parse::<usize>().ok()
            } else {
                None
            }
        } else {
            None
        };

        chunks.push(FortranFormatChunk::Data {
            descriptor,
            repeat: repeat_prefix.unwrap_or(1),
            width,
            precision,
        });
    }

    Some(chunks)
}

/// One item of a `print`/`write` list, as a string.
///
/// The rendering is ONE builtin that asks the value what it is, so a LOGICAL
/// writes `T`/`F` wherever it came from. This used to test the expression's
/// SHAPE here — a literal or a comparison became a `"True"`/`"False"` ternary,
/// anything else fell through to `__str__` and rendered `true` — which meant
/// `print *, .true.` and `print *, b` disagreed about the same type. The shape
/// test existed because a comparison stored into a LOGICAL was a raw i32 and
/// had no type left to ask; `materialize_bool_results` in the profile is what
/// makes it a boolean, and this is what reads it.
fn stringify_fortran_io_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__fortran_io_str")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn concat_fortran_io_parts(parts: Vec<Expression>) -> Expression {
    match parts.len() {
        0 => Expression::string(""),
        1 => parts
            .into_iter()
            .next()
            .unwrap_or_else(|| Expression::string("")),
        _ => Expression::new(ExprKind::Interpolation(
            parts
                .into_iter()
                .map(|part| match part.kind {
                    ExprKind::Lit(Literal::Str(text)) => InterpolPart::Text(text),
                    _ => InterpolPart::Expr(part),
                })
                .collect(),
        )),
    }
}

fn lower_fortran_implied_do_array_constructor(
    pair: Pair<Rule>,
) -> Result<Option<Expression>, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // The control variable splits the head values from the bounds. Reading a
    // fixed `parts[1]` only worked while the head was a single expression; the
    // head is a value LIST, so `((i*10 + j, j = 1, 2), i = 1, 3)` and
    // `(a, b, i = 1, 3)` both put the identifier further along.
    let Some(var_index) = parts.iter().position(|p| p.as_rule() == Rule::identifier) else {
        return Ok(None);
    };
    if var_index == 0 || parts.len() < var_index + 3 {
        return Ok(None);
    }

    // A head that is itself an implied-do contributes a RUN per iteration, not
    // one value, and so does a multi-value head — both make the mapped element
    // array-valued and need flattening afterwards.
    let mut head_values = Vec::new();
    let mut element_is_a_run = parts[..var_index].len() > 1;
    for head in &parts[..var_index] {
        if let Some(nested) = lower_fortran_implied_do_array_constructor(head.clone())? {
            element_is_a_run = true;
            head_values.push(vybe_ast::ArrayElement {
                key: None,
                value: nested,
                spread: true,
                by_ref: false,
            });
            continue;
        }
        let inner = if is_expr_rule(head.as_rule()) || head.as_rule() == Rule::expression {
            head.clone()
        } else {
            head.clone()
                .into_inner()
                .filter(|q| meaningful(q))
                .find(|q| is_expr_rule(q.as_rule()) || q.as_rule() == Rule::expression)
                .ok_or("empty implied-do value")?
        };
        head_values.push(vybe_ast::ArrayElement {
            key: None,
            value: walk_expr(inner)?,
            spread: false,
            by_ref: false,
        });
    }
    let element = if element_is_a_run {
        Expression::new(ExprKind::Array(head_values))
    } else {
        head_values
            .into_iter()
            .next()
            .ok_or("empty implied-do head")?
            .value
    };
    let loop_var = parts[var_index].as_str().to_string();
    let lower = walk_expr(parts[var_index + 1].clone())?;
    let upper = walk_expr(parts[var_index + 2].clone())?;
    let step = if let Some(step) = parts.get(var_index + 3) {
        walk_expr(step.clone())?
    } else {
        Expression::int(1)
    };

    let index_name = "__fortran_array_index";
    let size = build_fortran_implied_do_trip_count(lower.clone(), upper.clone(), step.clone());
    let current_value = build_fortran_implied_do_value(lower, step, index_name);
    let lowered_element = substitute_fortran_ident_expr(&element, &loop_var, &current_value);
    let array_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(size),
            Argument::positional(Expression::int(0)),
        ],
        optional: false,
    });

    let mapped = build_fortran_array_map(
        array_expr,
        lowered_element,
        true,
        "__fortran_array_item",
        index_name,
    );
    if !element_is_a_run {
        return Ok(Some(mapped));
    }
    Ok(Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(mapped),
            field: "flat".to_string(),
            null_safe: false,
        })),
        args: Vec::new(),
        optional: false,
    })))
}

fn build_fortran_implied_do_trip_count(
    lower: Expression,
    upper: Expression,
    step: Expression,
) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(upper),
                right: Box::new(lower),
            })),
            right: Box::new(step),
        })),
        right: Box::new(Expression::int(1)),
    })
}

fn build_fortran_implied_do_value(
    lower: Expression,
    step: Expression,
    index_name: &str,
) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(lower),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(Expression::ident(index_name)),
            right: Box::new(step),
        })),
    })
}

fn substitute_fortran_ident_expr(
    expr: &Expression,
    ident: &str,
    replacement: &Expression,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if name.eq_ignore_ascii_case(ident) => replacement.clone(),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(substitute_fortran_ident_expr(left, ident, replacement)),
            right: Box::new(substitute_fortran_ident_expr(right, ident, replacement)),
        }),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(substitute_fortran_ident_expr(inner, ident, replacement)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(substitute_fortran_ident_expr(cond, ident, replacement)),
            then: Box::new(substitute_fortran_ident_expr(then, ident, replacement)),
            else_: Box::new(substitute_fortran_ident_expr(else_, ident, replacement)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(substitute_fortran_ident_expr(object, ident, replacement)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(substitute_fortran_ident_expr(object, ident, replacement)),
            index: Box::new(substitute_fortran_ident_expr(index, ident, replacement)),
            null_safe: *null_safe,
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(substitute_fortran_ident_expr(callee, ident, replacement)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: substitute_fortran_ident_expr(&arg.value, ident, replacement),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(substitute_fortran_ident_expr(target, ident, replacement)),
            value: Box::new(substitute_fortran_ident_expr(value, ident, replacement)),
        }),
        ExprKind::Array(items) => Expression::new(ExprKind::Array(
            items
                .iter()
                .map(|item| vybe_ast::ArrayElement {
                    key: item
                        .key
                        .as_ref()
                        .map(|key| substitute_fortran_ident_expr(key, ident, replacement)),
                    value: substitute_fortran_ident_expr(&item.value, ident, replacement),
                    spread: item.spread,
                    by_ref: item.by_ref,
                })
                .collect(),
        )),
        ExprKind::Tuple(items) => Expression::new(ExprKind::Tuple(
            items
                .iter()
                .map(|item| substitute_fortran_ident_expr(item, ident, replacement))
                .collect(),
        )),
        ExprKind::Set(items) => Expression::new(ExprKind::Set(
            items
                .iter()
                .map(|item| substitute_fortran_ident_expr(item, ident, replacement))
                .collect(),
        )),
        ExprKind::Object(props) => Expression::new(ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: substitute_fortran_ident_expr(key, ident, replacement),
                        value: substitute_fortran_ident_expr(value, ident, replacement),
                    },
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: substitute_fortran_ident_expr(key, ident, replacement),
                        value: substitute_fortran_ident_expr(value, ident, replacement),
                    },
                    ObjectProperty::Spread(value) => ObjectProperty::Spread(
                        substitute_fortran_ident_expr(value, ident, replacement),
                    ),
                    _ => prop.clone(),
                })
                .collect(),
        )),
        ExprKind::Interpolation(parts) => Expression::new(ExprKind::Interpolation(
            parts
                .iter()
                .map(|part| match part {
                    InterpolPart::Expr(value) => {
                        InterpolPart::Expr(substitute_fortran_ident_expr(value, ident, replacement))
                    }
                    InterpolPart::Formatted(value, format) => InterpolPart::Formatted(
                        substitute_fortran_ident_expr(value, ident, replacement),
                        format.clone(),
                    ),
                    _ => part.clone(),
                })
                .collect(),
        )),
        ExprKind::IsType {
            expr: inner,
            type_name,
        } => Expression::new(ExprKind::IsType {
            expr: Box::new(substitute_fortran_ident_expr(inner, ident, replacement)),
            type_name: type_name.clone(),
        }),
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => Expression::new(ExprKind::Cast {
            expr: Box::new(substitute_fortran_ident_expr(inner, ident, replacement)),
            type_name: type_name.clone(),
        }),
        ExprKind::TypeOf(inner) => Expression::new(ExprKind::TypeOf(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::NullCoalesce { left, right } => Expression::new(ExprKind::NullCoalesce {
            left: Box::new(substitute_fortran_ident_expr(left, ident, replacement)),
            right: Box::new(substitute_fortran_ident_expr(right, ident, replacement)),
        }),
        ExprKind::Spread(inner) => Expression::new(ExprKind::Spread(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::Await(inner) => Expression::new(ExprKind::Await(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::Yield(Some(inner)) => Expression::new(ExprKind::Yield(Some(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        )))),
        ExprKind::YieldFrom(inner) => Expression::new(ExprKind::YieldFrom(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::SuperCall { method, args } => Expression::new(ExprKind::SuperCall {
            method: method.clone(),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: substitute_fortran_ident_expr(&arg.value, ident, replacement),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
        }),
        ExprKind::Slice { lower, upper, step } => Expression::new(ExprKind::Slice {
            lower: lower
                .as_ref()
                .map(|value| Box::new(substitute_fortran_ident_expr(value, ident, replacement))),
            upper: upper
                .as_ref()
                .map(|value| Box::new(substitute_fortran_ident_expr(value, ident, replacement))),
            step: step
                .as_ref()
                .map(|value| Box::new(substitute_fortran_ident_expr(value, ident, replacement))),
        }),
        ExprKind::Walrus { target, value } => Expression::new(ExprKind::Walrus {
            target: Box::new(substitute_fortran_ident_expr(target, ident, replacement)),
            value: Box::new(substitute_fortran_ident_expr(value, ident, replacement)),
        }),
        ExprKind::Void(inner) => Expression::new(ExprKind::Void(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::Delete(inner) => Expression::new(ExprKind::Delete(Box::new(
            substitute_fortran_ident_expr(inner, ident, replacement),
        ))),
        ExprKind::Sequence(items) => Expression::new(ExprKind::Sequence(
            items
                .iter()
                .map(|item| substitute_fortran_ident_expr(item, ident, replacement))
                .collect(),
        )),
        ExprKind::Range {
            start,
            end,
            inclusive,
        } => Expression::new(ExprKind::Range {
            start: Box::new(substitute_fortran_ident_expr(start, ident, replacement)),
            end: Box::new(substitute_fortran_ident_expr(end, ident, replacement)),
            inclusive: *inclusive,
        }),
        ExprKind::StaticAccess { class, member } => Expression::new(ExprKind::StaticAccess {
            class: Box::new(substitute_fortran_ident_expr(class, ident, replacement)),
            member: Box::new(substitute_fortran_ident_expr(member, ident, replacement)),
        }),
        _ => expr.clone(),
    }
}

fn walk_call(pair: Pair<Rule>) -> Result<Statement, String> {
    let inner = pair.into_inner().filter(|p| meaningful(p));
    let mut callee: Option<Expression> = None;
    let mut args = Vec::new();
    let mut alternate_returns = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::identifier | Rule::designator_name => {
                if p.as_str().eq_ignore_ascii_case("call") {
                    continue;
                }
                callee = Some(match callee.take() {
                    None => Expression::new(ExprKind::Ident(p.as_str().to_string())),
                    Some(expr) => Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    }),
                });
            }
            Rule::argument_list => {
                for a in p.into_inner() {
                    if a.as_rule() != Rule::argument {
                        continue;
                    }
                    // `*10` is not an argument at all — it is the label the
                    // callee's `return 1` picks, so it leaves the list and
                    // becomes a branch on the returned selector below.
                    if let Some(label) = fortran_alternate_return_label(&a) {
                        alternate_returns.push(label);
                        continue;
                    }
                    let (name, value) = walk_argument_expr(a)?;
                    args.push(Argument {
                        name,
                        value,
                        by_ref: false,
                        spread: false,
                    });
                }
            }
            _ => {}
        }
    }
    let expr = Expression::new(ExprKind::Call {
        callee: Box::new(callee.ok_or("missing call name")?),
        args,
        optional: false,
    });
    if !alternate_returns.is_empty() {
        return Ok(build_fortran_alternate_return_call(expr, alternate_returns));
    }
    let is_random_intrinsic = matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("random_number") || name.eq_ignore_ascii_case("random_seed"))
    );
    if !is_random_intrinsic {
        if let Some(stmt) = lower_intrinsic_statement(&expr) {
            return Ok(stmt);
        }
    }
    Ok(Statement::new(StmtKind::Expr(expr)))
}

/// `*10` in an argument list — the statement label an alternate return selects.
/// A bare `*` (assumed size) carries no label and is not one.
fn fortran_alternate_return_label(argument: &Pair<Rule>) -> Option<String> {
    let star = argument
        .clone()
        .into_inner()
        .find(|p| p.as_rule() == Rule::star_arg)?;
    fortran_first_statement_label(&star)
}

/// `call s(*10, *20)` — F77 alternate returns.
///
/// The callee's `return 1` / `return 2` names the Nth label rather than a
/// value, so the call becomes an ordinary call whose result selects a jump.
/// The jumps are the same `__fortran_goto` markers a written `goto` produces,
/// so the dispatch pass resolves them without knowing where they came from. A
/// plain `return` yields nothing, no arm matches, and control falls through to
/// the statement after the call — which is what the standard says.
fn build_fortran_alternate_return_call(call: Expression, labels: Vec<String>) -> Statement {
    let selector = "__fortran_alt_return";
    let mut labels = labels.into_iter().enumerate();
    let (_, first) = labels.next().expect("caller checked the list is not empty");
    let branch = Statement::new(StmtKind::If {
        cond: fortran_index_equals(&Expression::ident(selector), 1),
        then_body: vec![fortran_goto_marker_statement(first)],
        elifs: labels
            .map(|(index, label)| {
                (
                    fortran_index_equals(&Expression::ident(selector), index as i64 + 1),
                    vec![fortran_goto_marker_statement(label)],
                )
            })
            .collect(),
        else_body: None,
    });
    Statement::new(StmtKind::Block(vec![
        fortran_data_local(selector, call),
        branch,
    ]))
}

fn walk_argument_expr(pair: Pair<Rule>) -> Result<(Option<String>, Expression), String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    if inner.is_empty() {
        return Err("empty arg".to_string());
    }

    if inner.len() == 1 {
        return Ok((None, walk_argument_value(inner.pop().unwrap())?));
    }

    let first = inner.remove(0);
    if first.as_rule() == Rule::identifier && inner.len() == 1 {
        return Ok((
            Some(first.as_str().to_string()),
            walk_argument_value(inner.pop().unwrap())?,
        ));
    }

    Ok((None, walk_argument_value(inner.pop().unwrap())?))
}

fn walk_argument_value(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::slice_arg => walk_slice_arg(pair),
        // `operator(+)` as an ARGUMENT — F2018 `reduce`'s combining operation.
        // Carried as its symbol so the reduction below can pick the fold; there
        // is nothing to evaluate, and walking it as an expression produced a
        // call to an undefined `operator`.
        Rule::operator_arg => Ok(Expression::string(
            pair.as_str()
                .trim()
                .trim_start_matches(|c: char| c.is_alphabetic())
                .trim()
                .trim_start_matches('(')
                .trim_end_matches(')')
                .trim(),
        )),
        _ => walk_expr(pair),
    }
}

fn walk_slice_arg(pair: Pair<Rule>) -> Result<Expression, String> {
    let text = pair.as_str().trim();
    let segments: Vec<&str> = text.split(':').collect();
    let mut exprs = Vec::new();
    for part in pair.into_inner().filter(|p| meaningful(p)) {
        exprs.push(walk_expr(part)?);
    }
    let mut exprs = exprs.into_iter();

    let mut lower = None;
    let mut upper = None;
    let mut step = None;

    if segments
        .get(0)
        .is_some_and(|segment| !segment.trim().is_empty())
    {
        lower = exprs.next();
    }
    if segments
        .get(1)
        .is_some_and(|segment| !segment.trim().is_empty())
    {
        upper = exprs.next();
    }
    if segments
        .get(2)
        .is_some_and(|segment| !segment.trim().is_empty())
    {
        step = exprs.next();
    }

    Ok(Expression::new(ExprKind::Slice {
        lower: lower.map(Box::new),
        upper: upper.map(Box::new),
        step: step.map(Box::new),
    }))
}

fn walk_allocate_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    walk_allocate_stmt_from_text(pair.as_str())
}

fn walk_deallocate_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    walk_allocator_stmt(pair, "deallocate")
}

/// A bare derived-type name in an allocate type spec. An INTRINSIC spec
/// (`character(len=5)`, `integer`) names no constructor and is not one.
fn fortran_derived_type_spec_name(spec: &str) -> Option<String> {
    let spec = spec.trim();
    if !spec
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || c == '_')
        || spec.is_empty()
    {
        return None;
    }
    let lower = spec.to_ascii_lowercase();
    if matches!(
        lower.as_str(),
        "integer" | "real" | "double" | "complex" | "logical" | "character"
    ) {
        return None;
    }
    Some(spec.to_string())
}

/// `Child :: obj` → `("Child", "obj")`. Only the `::` that separates a type
/// spec from the item, so `obj(1:n)` is untouched.
fn split_fortran_alloc_type_spec(part: &str) -> Option<(&str, &str)> {
    let at = part.find("::")?;
    let (spec, rest) = part.split_at(at);
    let spec = spec.trim();
    if spec.is_empty() {
        return None;
    }
    Some((spec, rest[2..].trim()))
}

fn split_fortran_top_level_commas(text: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut paren = 0i32;
    let mut bracket = 0i32;
    let mut quote: Option<char> = None;
    let mut chars = text.char_indices().peekable();
    while let Some((idx, ch)) = chars.next() {
        if let Some(q) = quote {
            if ch == q {
                if chars.peek().is_some_and(|(_, next)| *next == q) {
                    chars.next();
                } else {
                    quote = None;
                }
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => paren += 1,
            ')' => paren -= 1,
            '[' => bracket += 1,
            ']' => bracket -= 1,
            ',' if paren == 0 && bracket == 0 => {
                parts.push(text[start..idx].trim());
                start = idx + ch.len_utf8();
            }
            _ => {}
        }
    }
    parts.push(text[start..].trim());
    parts.into_iter().filter(|part| !part.is_empty()).collect()
}

fn split_fortran_top_level_equals(text: &str) -> Option<(&str, &str)> {
    let mut paren = 0i32;
    let mut bracket = 0i32;
    let mut quote: Option<char> = None;
    let mut chars = text.char_indices().peekable();
    while let Some((idx, ch)) = chars.next() {
        if let Some(q) = quote {
            if ch == q {
                if chars.peek().is_some_and(|(_, next)| *next == q) {
                    chars.next();
                } else {
                    quote = None;
                }
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => paren += 1,
            ')' => paren -= 1,
            '[' => bracket += 1,
            ']' => bracket -= 1,
            '=' if paren == 0 && bracket == 0 => {
                return Some((text[..idx].trim(), text[idx + ch.len_utf8()..].trim()));
            }
            _ => {}
        }
    }
    None
}

fn parse_fortran_expression_text(text: &str) -> Result<Expression, String> {
    let mut pairs = FortranParser::parse(Rule::expression, text)
        .map_err(|err| format!("invalid Fortran expression `{text}`: {err}"))?;
    let pair = pairs
        .next()
        .ok_or_else(|| format!("missing Fortran expression `{text}`"))?;
    walk_expr(pair)
}

fn parse_fortran_alloc_target_text(
    text: &str,
    origins: &mut Vec<Option<Expression>>,
) -> Result<Expression, String> {
    let target_text = text
        .rsplit_once("::")
        .map(|(_, target)| target.trim())
        .unwrap_or_else(|| text.trim());
    let mut pairs = FortranParser::parse(Rule::alloc_item, target_text)
        .map_err(|err| format!("invalid allocate target `{target_text}`: {err}"))?;
    let pair = pairs
        .next()
        .ok_or_else(|| format!("missing allocate target `{target_text}`"))?;
    walk_alloc_item_expr_with_origins(pair, origins)
}

/// `__fortran_origin_v = [lo, …]` — the descriptor update that goes with an
/// `allocate` that stated an origin, so `lbound` answers what was allocated.
fn fortran_allocate_origin_statement(
    target: &Expression,
    origins: &[Option<Expression>],
) -> Option<Statement> {
    if !origins.iter().any(Option::is_some) {
        return None;
    }
    let ExprKind::Ident(name) = &fortran_allocation_target_place(target).kind else {
        return None;
    };
    Some(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&fortran_origin_variable_name(name))],
        value: Expression::new(ExprKind::Array(
            origins
                .iter()
                .map(|origin| ArrayElement {
                    key: None,
                    value: origin.clone().unwrap_or_else(|| Expression::int(1)),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )),
        by_ref: false,
    }))
}

fn fortran_allocation_target_place(target: &Expression) -> Expression {
    match &target.kind {
        ExprKind::Call { callee, .. } => callee.as_ref().clone(),
        _ => target.clone(),
    }
}

fn walk_allocate_stmt_from_text(text: &str) -> Result<Statement, String> {
    let Some(open) = text.find('(') else {
        return walk_allocator_stmt_text_fallback(text, "allocate");
    };
    let close = text
        .rfind(')')
        .ok_or_else(|| format!("missing allocate close paren in `{text}`"))?;
    let inner = &text[open + 1..close];
    let mut targets = Vec::new();
    let mut origin_statements = Vec::new();
    let mut source = None;
    let mut mold = None;
    let mut stat = None;
    let mut typed_constructions = Vec::new();
    let mut dynamic_type_markers = Vec::new();
    for part in split_fortran_top_level_commas(inner) {
        if let Some((name, value_text)) = split_fortran_top_level_equals(part) {
            if name.eq_ignore_ascii_case("source") {
                source = Some(parse_fortran_expression_text(value_text)?);
                continue;
            }
            if name.eq_ignore_ascii_case("mold") {
                mold = Some(parse_fortran_expression_text(value_text)?);
                continue;
            }
            if name.eq_ignore_ascii_case("stat") {
                stat = Some(parse_fortran_expression_text(value_text)?);
                continue;
            }
        }
        // `allocate(Child :: obj)` — a TYPED allocation names the dynamic type
        // the object is to have. The prefix was never read, so the object was
        // allocated with its DECLARED type and `select type` then took the
        // parent branch: `class is (Child)` never matched a `Child`.
        // ⛔ Only a DERIVED type names a constructor. `allocate(character(len=5)
        // :: s)` carries a type spec too, and treating that as a class emitted
        // `new character(len=5(` and dropped the allocation outright.
        let (part, typed_as, intrinsic_spec) = match split_fortran_alloc_type_spec(part) {
            Some((type_spec, rest)) => {
                let derived = fortran_derived_type_spec_name(type_spec);
                // An INTRINSIC spec names the dynamic type just as surely as a
                // derived one does; it just has no constructor to call. Kept so
                // `select type` can answer for it — see the marker below.
                let intrinsic = derived.is_none().then(|| type_spec.to_string());
                (rest, derived, intrinsic)
            }
            None => (part, None, None),
        };
        let mut origins = Vec::new();
        let target = parse_fortran_alloc_target_text(part, &mut origins)?;
        origin_statements.extend(fortran_allocate_origin_statement(&target, &origins));
        if let Some(spec) = intrinsic_spec {
            if let ExprKind::Ident(name) = &target.kind {
                // `allocate(integer :: val)` is the ONLY thing that gives an
                // unlimited polymorphic an intrinsic dynamic type, and the
                // runtime cannot recover it: an integer and a real are the same
                // f64 by the time `select type` asks. So the fact is carried
                // from here, where it is written, to the pass that needs it.
                dynamic_type_markers.push(Statement::new(StmtKind::Labeled {
                    label: format!(
                        "{FORTRAN_ALLOC_TYPE_MARKER}{}:{}",
                        name.to_ascii_lowercase(),
                        fortran_canonical_select_type_name(&spec)
                    ),
                    body: Box::new(Statement::new(StmtKind::Block(Vec::new()))),
                }));
            }
        }
        if let Some(type_name) = typed_as {
            // Constructing it IS the allocation: a derived type is built by
            // naming it, and that is what gives the object its dynamic type.
            typed_constructions.push(Statement::new(StmtKind::Assign {
                targets: vec![target.clone()],
                value: Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&type_name)),
                    args: Vec::new(),
                }),
                by_ref: false,
            }));
            continue;
        }
        targets.push(target);
    }

    let allocate_call = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("allocate")),
        args: targets
            .iter()
            .cloned()
            .map(Argument::positional)
        .collect::<Vec<_>>(),
        optional: false,
    })));

    if source.is_none()
        && mold.is_none()
        && stat.is_none()
        && origin_statements.is_empty()
        && typed_constructions.is_empty()
        && dynamic_type_markers.is_empty()
    {
        return Ok(allocate_call);
    }

    let mut statements = if targets.is_empty() {
        Vec::new()
    } else {
        vec![allocate_call]
    };
    statements.append(&mut dynamic_type_markers);
    statements.append(&mut typed_constructions);
    statements.append(&mut origin_statements);
    if let Some(source) = source {
        for target in &targets {
            statements.push(Statement::new(StmtKind::Assign {
                targets: vec![fortran_allocation_target_place(target)],
                value: source.clone(),
                by_ref: false,
            }));
        }
    } else if let Some(mold) = mold {
        for target in &targets {
            statements.push(Statement::new(StmtKind::Assign {
                targets: vec![fortran_allocation_target_place(target)],
                value: mold.clone(),
                by_ref: false,
            }));
        }
    }
    if let Some(stat) = stat {
        statements.push(Statement::new(StmtKind::Assign {
            targets: vec![stat],
            value: Expression::int(0),
            by_ref: false,
        }));
    }

    Ok(Statement::new(StmtKind::Block(statements)))
}

fn walk_allocator_stmt_text_fallback(text: &str, intrinsic_name: &str) -> Result<Statement, String> {
    let _ = text;
    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(intrinsic_name)),
        args: Vec::new(),
        optional: false,
    }))))
}

fn walk_allocator_stmt(pair: Pair<Rule>, intrinsic_name: &str) -> Result<Statement, String> {
    let mut args = Vec::new();
    for item in pair.into_inner().filter(|p| meaningful(p)) {
        match item.as_rule() {
            Rule::alloc_item => args.push(Argument::positional(walk_alloc_item_expr(item)?)),
            Rule::identifier => args.push(Argument::positional(Expression::ident(item.as_str()))),
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Call {
            callee: Box::new(Expression::ident(intrinsic_name)),
            args,
            optional: false,
        },
    ))))
}

fn walk_alloc_item_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    walk_alloc_item_expr_with_origins(pair, &mut Vec::new())
}

/// `allocate(v(-4:1))` states the array's origin as well as its shape, so the
/// declared lower bounds come back through `origins` for the caller to record.
fn walk_alloc_item_expr_with_origins(
    pair: Pair<Rule>,
    origins: &mut Vec<Option<Expression>>,
) -> Result<Expression, String> {
    let mut inner = pair.into_inner().filter(|p| meaningful(p));
    let ident = inner
        .next()
        .ok_or("missing allocate target")?
        .as_str()
        .to_string();
    let target = Expression::ident(&ident);

    let mut dims = Vec::new();
    for child in inner {
        if child.as_rule() != Rule::dimension_spec_list {
            continue;
        }
        for dim in child.into_inner().filter(|p| meaningful(p)) {
            let origin = walk_dimension_spec_origin(dim.clone())?;
            if let Some(expr) = walk_dimension_spec_expr(dim)? {
                origins.push(origin);
                dims.push(Argument::positional(expr));
            }
        }
    }

    if dims.is_empty() {
        Ok(target)
    } else {
        Ok(Expression::new(ExprKind::Call {
            callee: Box::new(target),
            args: dims,
            optional: false,
        }))
    }
}

/// The `lo` of a `lo:hi` dimension spec — `None` for a bare extent, which means
/// Fortran's default origin of 1.
fn walk_dimension_spec_origin(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    if pair.as_rule() != Rule::dimension_spec {
        return Ok(None);
    }
    let exprs: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|child| {
            meaningful(child) && (is_expr_rule(child.as_rule()) || child.as_rule() == Rule::expression)
        })
        .collect();
    if exprs.len() < 2 {
        return Ok(None);
    }
    walk_expr(exprs[0].clone()).map(Some)
}

fn walk_dimension_spec_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    match pair.as_rule() {
        Rule::dimension_spec => {
            let exprs: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|child| {
                    meaningful(child)
                        && (is_expr_rule(child.as_rule()) || child.as_rule() == Rule::expression)
                })
                .collect();
            // `allocate(v(-4:1))` states an origin and an upper bound, not an
            // extent. Taking the first expression would allocate `-4` elements.
            match exprs.len() {
                0 => Ok(None),
                1 => Ok(Some(walk_expr(exprs.into_iter().next().unwrap())?)),
                _ => {
                    let lo = walk_expr(exprs[0].clone())?;
                    let hi = walk_expr(exprs[1].clone())?;
                    Ok(Some(Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(hi),
                            right: Box::new(lo),
                        })),
                        right: Box::new(Expression::int(1)),
                    })))
                }
            }
        }
        rule if is_expr_rule(rule) || rule == Rule::expression => walk_expr(pair).map(Some),
        _ => Ok(None),
    }
}

fn walk_sub(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let modifiers = collect_fortran_proc_modifiers(&parts);
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        if p.as_rule() == Rule::identifier && nm.is_empty() {
            nm = p.as_str().to_string();
        } else if p.as_rule() == Rule::param_list {
            for pp in p.into_inner() {
                if pp.as_rule() == Rule::identifier {
                    params.push(Param {
                        name: pp.as_str().to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                }
            }
        } else {
            rest.push(p);
        }
    }
    apply_fortran_param_declaration_modes(&mut params, &rest);
    let mut body = walk_body(rest.into_iter())?;
    bind_fortran_param_declarations(&mut params, &mut body);
    promote_mutated_fortran_params(&mut params, &body);
    lower_fortran_namelist_io(&mut body);
    lower_fortran_body_intrinsics(&params, &mut body);
    lower_fortran_array_bounds(&params, &mut body);
    lower_fortran_array_semantics(&params, &mut body);
    let is_generator = body_has_yield(&body);
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm,
        params,
        return_type: None,
        body,
        modifiers,
        handles: vec![],
        is_async: false,
        is_generator,
        is_sub: true,
    }))
}

fn walk_func(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let modifiers = collect_fortran_proc_modifiers(&parts);
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rt = None;
    let mut result_name = None;
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::type_spec => {
                rt = Some(p.as_str().trim().to_string());
            }
            Rule::identifier => {
                if nm.is_empty() {
                    nm = p.as_str().to_string();
                }
            }
            Rule::param_list => {
                for pp in p.into_inner() {
                    if pp.as_rule() == Rule::identifier {
                        params.push(Param {
                            name: pp.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                }
            }
            Rule::proc_suffix | Rule::result_clause => {
                let suffix = p.as_str().trim().to_ascii_lowercase();
                if suffix.starts_with("result(") {
                    result_name = p
                        .into_inner()
                        .filter(|child| meaningful(child))
                        .find(|child| child.as_rule() == Rule::identifier)
                        .map(|child| child.as_str().to_string());
                }
            }
            _ => {
                rest.push(p);
            }
        }
    }
    apply_fortran_param_declaration_modes(&mut params, &rest);
    let mut body = walk_body(rest.into_iter())?;
    bind_fortran_param_declarations(&mut params, &mut body);
    promote_mutated_fortran_params(&mut params, &body);
    lower_fortran_namelist_io(&mut body);
    lower_fortran_body_intrinsics(&params, &mut body);
    normalize_fortran_function_result(&nm, result_name.as_deref(), &mut rt, &mut body);
    lower_fortran_array_bounds(&params, &mut body);
    lower_fortran_array_semantics(&params, &mut body);
    let is_generator = body_has_yield(&body);
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm,
        params,
        return_type: rt,
        body,
        modifiers,
        handles: vec![],
        is_async: false,
        is_generator,
        is_sub: false,
    }))
}

fn collect_fortran_proc_modifiers(parts: &[Pair<Rule>]) -> Modifiers {
    let mut modifiers = Modifiers::default();
    for part in parts {
        if part.as_rule() != Rule::proc_prefix {
            continue;
        }
        let name = match part.as_str().trim().to_ascii_lowercase().as_str() {
            "pure" => Some("pure"),
            "elemental" => Some("elemental"),
            "recursive" => Some("recursive"),
            "module" => Some("module"),
            "impure" => Some("impure"),
            "non_recursive" => Some("non_recursive"),
            _ => None,
        };
        if let Some(name) = name {
            modifiers.decorators.push(Expression::ident(name));
        }
    }
    modifiers
}

fn walk_interface_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let mut name: Option<String> = None;
    let mut members = Vec::new();

    for child in pair.into_inner().filter(|p| meaningful(p)) {
        let children: Vec<Pair<Rule>> = if child.as_rule() == Rule::statement_line {
            child.into_inner().filter(|p| meaningful(p)).collect()
        } else {
            vec![child]
        };

        for child in children {
            match child.as_rule() {
                Rule::identifier => {
                    if name.is_none() {
                        name = Some(child.as_str().to_string());
                    } else {
                        members.push(InterfaceMember::Method {
                            name: child.as_str().to_string(),
                            params: vec![],
                            return_type: None,
                            is_sub: true,
                            signature_source: None,
                        });
                    }
                }
                Rule::interface_designator => {
                    if name.is_none() {
                        let designator = child.as_str().trim();
                        let normalized = if designator.starts_with("operator(")
                            || designator.starts_with("assignment(")
                            || designator.starts_with("read(")
                            || designator.starts_with("write(")
                        {
                            designator.to_string()
                        } else if matches!(designator, "+" | "-" | "*" | "/" | "**") {
                            format!("operator({designator})")
                        } else if designator == "=" {
                            "assignment(=)".to_string()
                        } else {
                            designator.to_string()
                        };
                        name = Some(normalized);
                    }
                }
                Rule::subroutine_decl | Rule::function_decl => {
                    members.push(walk_interface_member(child)?);
                }
                Rule::procedure_decl => {
                    members.extend(walk_interface_procedure_members(child)?);
                }
                Rule::interface_end => break,
                _ => {}
            }
        }
    }

    if members.is_empty() {
        return Ok(Some(Statement::new(StmtKind::Empty)));
    }

    let interface_name = name.unwrap_or_else(|| {
        members
            .first()
            .and_then(|member| match member {
                InterfaceMember::Method { name, .. } => Some(name.clone()),
                _ => None,
            })
            .unwrap_or_else(|| "__fortran_interface".to_string())
    });

    Ok(Some(Statement::new(StmtKind::InterfaceDecl {
        name: interface_name,
        parents: vec![],
        members,
        decorators: vec![],
    })))
}

fn walk_interface_member(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let decl = match pair.as_rule() {
        Rule::subroutine_decl => walk_sub(pair)?,
        Rule::function_decl => walk_func(pair)?,
        _ => return Err("unsupported Fortran interface member".to_string()),
    };

    match decl.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            is_sub,
            ..
        } => Ok(InterfaceMember::Method {
            name,
            params,
            return_type,
            is_sub,
            signature_source: None,
        }),
        _ => Err("expected function declaration in Fortran interface".to_string()),
    }
}

fn walk_interface_procedure_members(pair: Pair<Rule>) -> Result<Vec<InterfaceMember>, String> {
    let decl = walk_procedure_decl(pair)?;
    let StmtKind::VarDecl { declarations, .. } = decl.kind else {
        return Err("expected procedure declaration in Fortran interface".to_string());
    };

    let mut members = Vec::new();
    for declaration in declarations {
        let BindingPattern::Ident(name) = declaration.pattern else {
            continue;
        };

        let signature_source = declaration.type_hint.as_deref().and_then(|type_hint| {
            type_hint
                .strip_prefix("procedure(")
                .and_then(|rest| rest.strip_suffix(')'))
                .map(|source| source.trim().to_string())
        });

        members.push(InterfaceMember::Method {
            name,
            params: vec![],
            return_type: None,
            is_sub: true,
            signature_source,
        });
    }

    Ok(members)
}

fn walk_type(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut nm = String::new();
    let mut members = Vec::new();
    let mut parents = Vec::new();
    let mut modifiers = ClassModifiers::default();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::identifier => {
                if nm.is_empty() {
                    nm = p.as_str().to_string();
                }
            }
            Rule::type_attribute => {
                apply_fortran_type_attribute(p, &mut modifiers, &mut parents);
            }
            Rule::type_member => {
                for m in p.into_inner() {
                    if matches!(m.as_rule(), Rule::var_declaration | Rule::procedure_decl) {
                        let decl = match m.as_rule() {
                            Rule::var_declaration => walk_var_decl(m)?,
                            Rule::procedure_decl => walk_procedure_decl(m)?,
                            _ => unreachable!(),
                        };
                        if let StmtKind::VarDecl { declarations, .. } = &decl.kind {
                            for d in declarations {
                                if let BindingPattern::Ident(fname) = &d.pattern {
                                    let field_type_hint = d.type_hint.as_ref().map(|type_hint| {
                                        fortran_array_type_hint(
                                            type_hint,
                                            d.array_bounds.as_deref(),
                                        )
                                    });
                                    members.push(ClassMember::Field {
                                        name: fname.clone(),
                                        type_hint: field_type_hint,
                                        init: d.init.clone(),
                                        modifiers: Modifiers::default(),
                                        with_events: false,
                                        array_bounds: d.array_bounds.clone(),
                                        storage: None,
                                    });
                                }
                            }
                        }
                    }
                }
            }
            Rule::type_bound_procedure => {
                let raw_rule = p.as_str().trim().to_ascii_lowercase();
                let is_generic_binding = raw_rule.starts_with("generic");
                let is_final_binding = raw_rule.starts_with("final");
                let mut method_bindings = Vec::new();
                let mut method_modifiers = Modifiers::default();
                let mut generic_or_final_names = Vec::new();
                for child in p.into_inner().filter(|p| meaningful(p)) {
                    match child.as_rule() {
                        Rule::tbp_attribute => {
                            apply_fortran_type_bound_attribute(
                                child.as_str(),
                                &mut method_modifiers,
                            );
                        }
                        Rule::tbp_binding => {
                            if let Some((public_name, implementation_name)) =
                                parse_tbp_binding(child)
                            {
                                method_bindings.push((public_name, implementation_name));
                            }
                        }
                        // `generic :: operator(+) => add` — the bound name is a
                        // DESIGNATOR now, so a plain `generic :: g => impl`
                        // arrives wrapped and stopped being seen at all: the
                        // implementation name became the public one.
                        Rule::interface_designator if is_generic_binding => {
                            generic_or_final_names.push(child.as_str().trim().to_string());
                        }
                        Rule::identifier | Rule::designator_name
                            if is_generic_binding || is_final_binding =>
                        {
                            generic_or_final_names.push(child.as_str().to_string());
                        }
                        _ => {}
                    }
                }

                if method_bindings.is_empty() {
                    if is_generic_binding {
                        if let Some((public_name, implementation_names)) =
                            generic_or_final_names.split_first()
                        {
                            for implementation_name in implementation_names {
                                method_bindings
                                    .push((public_name.clone(), implementation_name.clone()));
                            }
                        }
                    } else {
                        for name in generic_or_final_names {
                            method_bindings.push((name.clone(), name));
                        }
                    }
                }

                method_bindings.dedup_by(|left, right| {
                    left.0.eq_ignore_ascii_case(&right.0) && left.1.eq_ignore_ascii_case(&right.1)
                });
                for (method_name, implementation_name) in method_bindings {
                    members.push(ClassMember::Method(Box::new(Statement::new(
                        StmtKind::FunctionDecl {
                            name: method_name,
                            params: vec![],
                            return_type: None,
                            body: vec![],
                            modifiers: method_modifiers.clone(),
                            handles: vec![type_bound_impl_handle(&implementation_name)],
                            is_async: false,
                            is_generator: false,
                            is_sub: true,
                        },
                    ))));
                }
            }
            _ => {}
        }
    }
    Ok(Statement::new(StmtKind::ClassDecl {
        name: nm,
        parents,
        interfaces: vec![],
        members,
        modifiers: ClassModifiers {
            // A Fortran derived type is a VALUE aggregate: intrinsic assignment
            // is component-wise. Measured with gfortran — `b = a; b%x = 99`
            // leaves `a%x` at 1.
            //
            // Equality stays Identity: the standard gives a derived type no
            // intrinsic `==`; a program must define one via an interface.
            semantics: ValueSemantics {
                storage: ValueStorage::Value,
                ..Default::default()
            },
            ..modifiers
        },
        decorators: vec![],
    }))
}

fn parse_tbp_binding(pair: Pair<Rule>) -> Option<(String, String)> {
    let mut names = pair
        .into_inner()
        .filter(|binding_part| {
            matches!(
                binding_part.as_rule(),
                Rule::identifier | Rule::designator_name
            )
        })
        .map(|binding_part| binding_part.as_str().to_string());
    let public_name = names.next()?;
    let implementation_name = names.next().unwrap_or_else(|| public_name.clone());
    Some((public_name, implementation_name))
}

fn type_bound_impl_handle(name: &str) -> String {
    format!("{FORTRAN_TBP_IMPL_HANDLE_PREFIX}{name}")
}

fn type_bound_impl_name(handles: &[String]) -> Option<&str> {
    handles
        .iter()
        .find_map(|handle| handle.strip_prefix(FORTRAN_TBP_IMPL_HANDLE_PREFIX))
}

fn collect_global_procedures(body: &[Statement]) -> HashMap<String, Vec<Statement>> {
    let mut pool: HashMap<String, Vec<Statement>> = HashMap::new();
    for stmt in body.iter() {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } => {
                pool.entry(name.to_ascii_lowercase())
                    .or_default()
                    .push(stmt.clone());
            }
            StmtKind::ModuleDecl { members, .. } => {
                for m in members.iter() {
                    if let ClassMember::Method(s) = m {
                        if let StmtKind::FunctionDecl { name, .. } = &s.kind {
                            pool.entry(name.to_ascii_lowercase())
                                .or_default()
                                .push((**s).clone());
                        }
                    }
                }
            }
            _ => {}
        }
    }
    pool
}

fn bind_module_type_bound_procedures_with_pool(
    members: &mut [ClassMember],
    extra_pool: &HashMap<String, Vec<Statement>>,
) {
    let mut procedures: HashMap<String, Vec<Statement>> = extra_pool.clone();
    let mut bound_impl_targets: Vec<(String, String)> = Vec::new();
    for member in members.iter() {
        if let ClassMember::Method(stmt) = member {
            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                procedures
                    .entry(name.to_ascii_lowercase())
                    .or_default()
                    .push((**stmt).clone());
            }
        }
    }

    for member in members.iter_mut() {
        let ClassMember::NestedType(stmt) = member else {
            continue;
        };
        let StmtKind::ClassDecl {
            name,
            members: class_members,
            ..
        } = &mut stmt.kind
        else {
            continue;
        };

        for class_member in class_members.iter_mut() {
            let ClassMember::Method(method_stmt) = class_member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                body,
                modifiers,
                handles,
                ..
            } = &method_stmt.kind
            else {
                continue;
            };

            if !body.is_empty() {
                continue;
            }

            let implementation_name = type_bound_impl_name(handles).unwrap_or(method_name);
            let Some(candidates) = procedures.get(&implementation_name.to_ascii_lowercase()) else {
                continue;
            };
            // A `nopass` binding names a procedure with no receiver, so
            // requiring one of the type's own dummies never matches and the
            // method was left with an EMPTY body — `call obj%s()` ran nothing.
            let nopass = modifiers.is_static;
            let Some(candidate) = candidates
                .iter()
                .find(|candidate| nopass || function_decl_targets_type(candidate, name))
            else {
                continue;
            };
            bound_impl_targets.push((implementation_name.to_ascii_lowercase(), name.clone()));
            let StmtKind::FunctionDecl {
                params,
                return_type,
                body,
                is_async,
                is_generator,
                is_sub,
                ..
            } = &candidate.kind
            else {
                continue;
            };

            *method_stmt = Box::new(Statement {
                kind: StmtKind::FunctionDecl {
                    name: method_name.clone(),
                    params: params.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: Vec::new(),
                    is_async: *is_async,
                    is_generator: *is_generator,
                    is_sub: *is_sub,
                },
                span: method_stmt.span,
            });

            if let StmtKind::FunctionDecl { params, .. } = &mut method_stmt.kind {
                if let Some(first_param) = params.first_mut() {
                    first_param.pass_by = PassBy::Value;
                }
            }
        }
    }

    for member in members.iter_mut() {
        let ClassMember::Method(stmt) = member else {
            continue;
        };

        let should_demote_self_ref = {
            let StmtKind::FunctionDecl {
                name, params, body, ..
            } = &stmt.kind
            else {
                continue;
            };
            bound_impl_targets
                .iter()
                .any(|(implementation_name, class_name)| {
                    name.eq_ignore_ascii_case(implementation_name)
                        && function_decl_targets_type_parts(params, body, class_name)
                })
        };

        if !should_demote_self_ref {
            continue;
        }

        let StmtKind::FunctionDecl { params, .. } = &mut stmt.kind else {
            continue;
        };
        if let Some(first_param) = params.first_mut() {
            first_param.pass_by = PassBy::Value;
        }
    }
}

fn bind_top_level_type_bound_procedures(body: &mut [Statement]) {
    // Collect procedures from everywhere so cross-module TBP binding works.
    let global_pool = collect_global_procedures(body);
    let procedures = global_pool.clone();
    let mut bound_impl_targets: Vec<(String, String)> = Vec::new();

    for stmt in body.iter_mut() {
        let (name, members) = match &mut stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => (name.clone(), members),
            StmtKind::ModuleDecl { members, .. } => {
                bind_module_type_bound_procedures_with_pool(members, &global_pool);
                continue;
            }
            _ => continue,
        };

        for member in members.iter_mut() {
            let ClassMember::Method(method_stmt) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                body,
                modifiers,
                handles,
                ..
            } = &method_stmt.kind
            else {
                continue;
            };

            if !body.is_empty() {
                continue;
            }

            let implementation_name = type_bound_impl_name(handles).unwrap_or(method_name);
            let Some(candidates) = procedures.get(&implementation_name.to_ascii_lowercase()) else {
                continue;
            };
            let Some(candidate) = candidates
                .iter()
                .find(|candidate| function_decl_targets_type(candidate, &name))
            else {
                continue;
            };
            bound_impl_targets.push((implementation_name.to_ascii_lowercase(), name.clone()));
            let StmtKind::FunctionDecl {
                params,
                return_type,
                body,
                is_async,
                is_generator,
                is_sub,
                ..
            } = &candidate.kind
            else {
                continue;
            };

            *method_stmt = Box::new(Statement {
                kind: StmtKind::FunctionDecl {
                    name: method_name.clone(),
                    params: params.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: Vec::new(),
                    is_async: *is_async,
                    is_generator: *is_generator,
                    is_sub: *is_sub,
                },
                span: method_stmt.span,
            });

            if let StmtKind::FunctionDecl { params, .. } = &mut method_stmt.kind {
                if let Some(first_param) = params.first_mut() {
                    first_param.pass_by = PassBy::Value;
                }
            }
        }
    }

    for stmt in body.iter_mut() {
        let should_demote_self_ref = {
            let StmtKind::FunctionDecl {
                name, params, body, ..
            } = &stmt.kind
            else {
                continue;
            };
            bound_impl_targets
                .iter()
                .any(|(implementation_name, class_name)| {
                    name.eq_ignore_ascii_case(implementation_name)
                        && function_decl_targets_type_parts(params, body, class_name)
                })
        };

        if !should_demote_self_ref {
            continue;
        }

        let StmtKind::FunctionDecl { params, .. } = &mut stmt.kind else {
            continue;
        };
        if let Some(first_param) = params.first_mut() {
            first_param.pass_by = PassBy::Value;
        }
    }
}

fn function_decl_targets_type(stmt: &Statement, class_name: &str) -> bool {
    let StmtKind::FunctionDecl { params, body, .. } = &stmt.kind else {
        return false;
    };
    function_decl_targets_type_parts(params, body, class_name)
}

fn function_decl_targets_type_parts(
    params: &[Param],
    body: &[Statement],
    class_name: &str,
) -> bool {
    let Some(first_param) = params.first() else {
        return false;
    };

    if first_param
        .type_hint
        .as_deref()
        .and_then(parse_derived_type_name)
        .is_some_and(|target_type| target_type.eq_ignore_ascii_case(class_name))
    {
        return true;
    }

    body.iter()
        .find_map(|statement| {
            let StmtKind::VarDecl { declarations, .. } = &statement.kind else {
                return None;
            };
            declarations.iter().find_map(|declaration| {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    return None;
                };
                if !name.eq_ignore_ascii_case(&first_param.name) {
                    return None;
                }
                declaration
                    .type_hint
                    .as_deref()
                    .and_then(parse_derived_type_name)
            })
        })
        .is_some_and(|target_type| target_type.eq_ignore_ascii_case(class_name))
}

// ── Expressions ────────────────────────────────────────────────────────────

fn walk_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::expression
        | Rule::logical_equiv
        | Rule::logical_or
        | Rule::logical_and
        | Rule::logical_not
        | Rule::comparison
        | Rule::addition
        | Rule::multiplication
        | Rule::power
        | Rule::concat
        | Rule::unary => walk_binop(pair),
        Rule::primary_expr => {
            // primary_atom followed by zero or more postfix_op (member, call, coindex).
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let atom = inner.next().ok_or("empty primary")?;
            let mut expr = walk_expr(atom)?;
            for op in inner {
                if op.as_rule() != Rule::postfix_op {
                    continue;
                }
                let mut op_inner = op.clone().into_inner().filter(|p| meaningful(p));
                let first = op_inner.next();
                match first {
                    Some(p) if matches!(p.as_rule(), Rule::identifier | Rule::designator_name) => {
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
                                let (_, value) = walk_argument_expr(a)?;
                                args.push(Argument::positional(value));
                            }
                        }
                        if let Some(lowered) = lower_intrinsic_expr_call(&expr, &args) {
                            expr = lowered;
                        } else {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args,
                                optional: false,
                            });
                        }
                    }
                    Some(_) => { /* ignore */ }
                    None => {
                        // Empty `()` — call with no args
                        if let Some(lowered) = lower_intrinsic_expr_call(&expr, &[]) {
                            expr = lowered;
                        } else {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args: Vec::new(),
                                optional: false,
                            });
                        }
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
            // `[a, b, c]` / `(/ a, b, c /)` / implied-do.
            // Pure implied-do constructors lower to a shared Array.map
            // so dynamic arrays materialize in the common AST.
            let mut values = Vec::new();
            for p in pair.clone().into_inner() {
                if matches!(p.as_rule(), Rule::array_constructor_body) {
                    for v in p.into_inner().filter(|q| meaningful(q)) {
                        if matches!(v.as_rule(), Rule::array_constructor_value) {
                            values.push(v);
                        }
                    }
                }
            }
            if values.len() == 1 {
                if let Some(lowered) =
                    lower_fortran_implied_do_array_constructor(values[0].clone())?
                {
                    return Ok(lowered);
                }
            }

            let mut elems: Vec<vybe_ast::ArrayElement> = Vec::new();
            for value in values {
                // An implied-do among other values contributes ITS WHOLE RUN,
                // so it spreads. Taking the first inner expression — which is
                // what the plain path below does — reduced `(10, i = 1, 2)` to
                // the single element `10`, so `(/ (10,i=1,2), (20,i=1,3) /)`
                // built a 2-element array instead of a 5-element one.
                if let Some(lowered) = lower_fortran_implied_do_array_constructor(value.clone())? {
                    elems.push(vybe_ast::ArrayElement {
                        key: None,
                        value: lowered,
                        spread: true,
                        by_ref: false,
                    });
                    continue;
                }
                if let Some(inner) = value
                    .into_inner()
                    .filter(|q| meaningful(q))
                    .find(|q| is_expr_rule(q.as_rule()) || matches!(q.as_rule(), Rule::expression))
                {
                    elems.push(vybe_ast::ArrayElement {
                        key: None,
                        value: walk_expr(inner)?,
                        spread: false,
                        by_ref: false,
                    });
                }
            }
            Ok(Expression::new(ExprKind::Array(elems)))
        }
        Rule::literal => walk_expr(pair.into_inner().next().ok_or("empty literal")?),
        Rule::logical_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(
            pair.as_str().to_lowercase().contains("true"),
        )))),
        Rule::number_literal => {
            let s = pair.as_str().trim();
            let clean = s.split('_').next().unwrap_or(s);
            if clean.contains('.')
                || clean.to_lowercase().contains('e')
                || clean.to_lowercase().contains('d')
            {
                let n: f64 = clean
                    .replace('d', "e")
                    .replace('D', "E")
                    .parse()
                    .unwrap_or(0.0);
                Ok(Expression::new(ExprKind::Lit(Literal::Float(n))))
            } else {
                let n: i64 = clean.parse().unwrap_or(0);
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            }
        }
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(
            parse_fortran_string_literal_text(pair.as_str()),
        )))),
        Rule::boz_literal => {
            // `b'..'` / `o'..'` / `z'..'` — bit / octal / hex literal.
            let s = pair.as_str();
            let prefix = s.chars().next().unwrap_or('z').to_ascii_lowercase();
            let body = &s[1..];
            let trimmed = body.trim_matches(|c: char| c == '\'' || c == '"');
            let radix = match prefix {
                'b' => 2,
                'o' => 8,
                _ => 16,
            };
            let n = i64::from_str_radix(trimmed, radix).unwrap_or(0);
            Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
        }
        Rule::identifier => Ok(fortran_iso_c_binding_constant(pair.as_str())
            .unwrap_or_else(|| Expression::new(ExprKind::Ident(pair.as_str().to_string())))),
        Rule::designator_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::function_call_or_subscript => {
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let nm = inner.next().ok_or("missing fn")?.as_str().to_string();
            let mut args = Vec::new();
            // The literal's SPELLING, kept only long enough to answer `kind`.
            // `1.0d0` and `1.0` both walk to `Literal::Float`, so by the time
            // the AST exists the `d` exponent and any `_8` suffix are gone and
            // nothing downstream can tell a double literal from a default one.
            let mut arg_texts = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for a in p.into_inner() {
                        if a.as_rule() == Rule::argument {
                            arg_texts.push(a.as_str().trim().to_string());
                            // KEEP the keyword name. `open(unit=7,
                            // status='scratch')` has no dedicated grammar rule
                            // and arrives here, and dropping the names left
                            // every `open` specifier identified by POSITION
                            // alone — which is why `status='scratch'` was being
                            // used as the FILENAME.
                            let (name, value) = walk_argument_expr(a)?;
                            args.push(Argument {
                                value,
                                name,
                                by_ref: false,
                                spread: false,
                            });
                        }
                    }
                }
            }
            if nm.eq_ignore_ascii_case("kind") && arg_texts.len() == 1 {
                if let Some(kind) = fortran_literal_kind_from_text(&arg_texts[0]) {
                    return Ok(Expression::int(kind));
                }
            }
            // `out_of_range` reads its MOLD's kind from the literal suffix, so
            // it has to fold here while `arg_texts` still exists.
            if nm.eq_ignore_ascii_case("out_of_range") && args.len() >= 2 && arg_texts.len() >= 2 {
                let round = args
                    .iter()
                    .find(|a| {
                        a.name
                            .as_deref()
                            .is_some_and(|n| n.eq_ignore_ascii_case("round"))
                    })
                    .map(|a| a.value.clone())
                    .or_else(|| args.get(2).map(|a| a.value.clone()));
                if let Some(folded) = build_fortran_out_of_range_expr(
                    args[0].value.clone(),
                    &arg_texts[1],
                    round,
                ) {
                    return Ok(folded);
                }
            }
            let callee = Expression::new(ExprKind::Ident(nm));
            if let Some(lowered) = lower_intrinsic_expr_call(&callee, &args) {
                Ok(lowered)
            } else {
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(callee),
                    args,
                    optional: false,
                }))
            }
        }
        Rule::argument => Ok(walk_argument_expr(pair)?.1),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn lower_intrinsic_statement(expr: &Expression) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    // ── Single-image collectives ────────────────────────────────────────────
    //
    // One image is the whole team, so a collective has nothing to combine: the
    // value it would reduce ACROSS images is already the answer, and gfortran
    // under `-fcoarray=single` leaves the argument untouched. An empty block is
    // the statement that says so.
    if matches!(
        name.to_ascii_lowercase().as_str(),
        "co_sum"
            | "co_min"
            | "co_max"
            | "co_product"
            | "co_reduce"
            | "co_broadcast"
            | "random_init"
            | "event_post"
            | "event_wait"
            // Team management on one image: there is only ever the initial
            // team, so forming, changing and syncing one are all no-ops.
            | "form_team"
            | "change_team"
            | "end_team"
            | "sync_team"
    ) {
        return Some(Statement::new(StmtKind::Block(Vec::new())));
    }

    // `event_query(event, count)` — nothing has been posted on one image.
    if name.eq_ignore_ascii_case("event_query") && args.len() >= 2 {
        return Some(Statement::new(StmtKind::Assign {
            targets: vec![args[1].value.clone()],
            value: Expression::int(0),
            by_ref: false,
        }));
    }

    // `atomic_define(atom, value)` / `atomic_ref(value, atom)` — on one image
    // there is nothing to serialise against, so the atomic pair IS the
    // assignment, in the argument order each one declares.
    if let Some((target, source)) = match name.to_ascii_lowercase().as_str() {
        "atomic_define" if args.len() >= 2 => Some((0, 1)),
        "atomic_ref" if args.len() >= 2 => Some((0, 1)),
        _ => None,
    } {
        return Some(Statement::new(StmtKind::Assign {
            targets: vec![args[target].value.clone()],
            value: args[source].value.clone(),
            by_ref: false,
        }));
    }

    if name.eq_ignore_ascii_case("nullify") {
        if args.is_empty() {
            return None;
        }

        let assigns = args
            .iter()
            .map(|arg| {
                Statement::new(StmtKind::Assign {
                    targets: vec![arg.value.clone()],
                    value: Expression::null(),
                    by_ref: false,
                })
            })
            .collect::<Vec<_>>();
        return Some(Statement::new(StmtKind::Block(assigns)));
    }

    if name.eq_ignore_ascii_case("move_alloc") && args.len() >= 2 {
        let from = args[0].value.clone();
        let to = args[1].value.clone();
        return Some(Statement::new(StmtKind::Block(vec![
            Statement::new(StmtKind::Assign {
                targets: vec![to],
                value: from.clone(),
                by_ref: false,
            }),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("deallocate")),
                args: vec![Argument::positional(from)],
                optional: false,
            }))),
        ])));
    }

    if name.eq_ignore_ascii_case("open") {
        return lower_open_intrinsic_statement(expr, args);
    }

    if name.eq_ignore_ascii_case("close") {
        return Some(Statement::new(StmtKind::CloseFile(
            args.iter()
                .find(|arg| {
                    arg.name
                        .as_deref()
                        .is_none_or(|name| name.eq_ignore_ascii_case("unit"))
                })
                .map(|arg| arg.value.clone()),
        )));
    }

    if name.eq_ignore_ascii_case("inquire") {
        return lower_fortran_inquire_statement(args);
    }

    // Our records are written through on each write rather than buffered in the
    // C sense, so there is nothing pending for `flush` to push. `endfile` marks
    // the end of the record sequence, which is where the file already ends.
    if name.eq_ignore_ascii_case("flush") || name.eq_ignore_ascii_case("endfile") {
        return Some(Statement::new(StmtKind::Block(Vec::new())));
    }

    if name.eq_ignore_ascii_case("rewind") {
        return Some(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__fortran_rewind")),
                args: args.to_vec(),
                optional: false,
            },
        ))));
    }

    if name.eq_ignore_ascii_case("random_number") {
        return lower_fortran_random_number_statement(args);
    }

    if name.eq_ignore_ascii_case("random_seed") {
        return Some(lower_fortran_random_seed_statement(args));
    }

    None
}

/// `INQUIRE(file=…, exist=…, size=…)` — ask the file system about a file.
///
/// The properties that name a FILE are answered from the same `wasi:filesystem`
/// surface python and php bind (`exists`, `fileSize`); a keyword is only
/// emitted when it can be answered truthfully.
///
/// ⛔ The unit-based properties — `opened=`, `number=`, `form=`, `access=`,
/// `name=` — describe the compiler's unit table, which has no query surface, so
/// they are LEFT ALONE rather than filled with a plausible guess. A program
/// asking for them keeps whatever its variable held; it does not get a lie.
fn lower_fortran_inquire_statement(args: &[Argument]) -> Option<Statement> {
    let keyword = |want: &str| {
        args.iter()
            .find(|arg| {
                arg.name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case(want))
            })
            .map(|arg| arg.value.clone())
    };
    let file = keyword("file")?;
    let mut body = Vec::new();
    for (want, builtin) in [
        ("exist", "__fortran_file_exists"),
        ("size", "__fortran_file_size"),
    ] {
        let Some(target) = keyword(want) else {
            continue;
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(builtin)),
                args: vec![Argument::positional(file.clone())],
                optional: false,
            }),
            by_ref: false,
        }));
    }
    (!body.is_empty()).then(|| Statement::new(StmtKind::Block(body)))
}

fn lower_fortran_random_number_statement(args: &[Argument]) -> Option<Statement> {
    let target = args
        .iter()
        .find(|arg| {
            arg.name
                .as_deref()
                .is_none_or(|name| name.eq_ignore_ascii_case("harvest"))
        })?
        .value
        .clone();
    Some(Statement::new(StmtKind::Assign {
        targets: vec![fortran_random_assignment_target(target.clone())],
        value: fortran_random_value_for_target(target),
        by_ref: false,
    }))
}

fn lower_fortran_random_seed_statement(args: &[Argument]) -> Statement {
    let seed_store = Expression::ident("__vybe_fortran_random_seed");
    if args.is_empty() {
        return Statement::new(StmtKind::Assign {
            targets: vec![seed_store],
            value: Expression::new(ExprKind::Array(vec![ArrayElement {
                key: None,
                value: Expression::int(1),
                spread: false,
                by_ref: false,
            }])),
            by_ref: false,
        });
    }

    let mut statements = Vec::new();
    for arg in args {
        match arg.name.as_deref().map(str::to_ascii_lowercase).as_deref() {
            Some("size") => statements.push(Statement::new(StmtKind::Assign {
                targets: vec![arg.value.clone()],
                value: Expression::int(8),
                by_ref: false,
            })),
            Some("put") => statements.push(Statement::new(StmtKind::Assign {
                targets: vec![seed_store.clone()],
                value: arg.value.clone(),
                by_ref: false,
            })),
            Some("get") => statements.push(Statement::new(StmtKind::Assign {
                targets: vec![arg.value.clone()],
                value: seed_store.clone(),
                by_ref: false,
            })),
            _ => {}
        }
    }

    if statements.is_empty() {
        Statement::new(StmtKind::Block(vec![]))
    } else if statements.len() == 1 {
        statements
            .pop()
            .unwrap_or_else(|| Statement::new(StmtKind::Block(vec![])))
    } else {
        Statement::new(StmtKind::Block(statements))
    }
}

fn fortran_random_assignment_target(target: Expression) -> Expression {
    match target.kind {
        ExprKind::Index { object, .. } => *object,
        _ => target,
    }
}

fn fortran_random_target_is_array(target: &Expression, type_env: &HashMap<String, String>) -> bool {
    match &target.kind {
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .is_some_and(|hint| hint.trim_end().ends_with("()")),
        ExprKind::Member { field, .. } => type_env
            .get(&field.to_ascii_lowercase())
            .is_some_and(|hint| hint.trim_end().ends_with("()")),
        ExprKind::Index { .. } => true,
        _ => false,
    }
}

fn fortran_random_array_fill_expr(target: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(target),
            field: "map".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params: vec![
                Param {
                    name: "__fortran_random_item".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
                Param {
                    name: "__fortran_random_index".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
            ],
            body: LambdaBody::Expr(Box::new(Expression::float(0.5))),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    })
}

fn fortran_random_value_for_target(target: Expression) -> Expression {
    if matches!(target.kind, ExprKind::Ident(_) | ExprKind::Index { .. }) {
        return Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("Array")),
                    field: "isArray".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(fortran_random_assignment_target(
                    target.clone(),
                ))],
                optional: false,
            })),
            then: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(fortran_random_assignment_target(target.clone())),
                    field: "map".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
                    params: vec![
                        Param {
                            name: "__fortran_random_item".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                        Param {
                            name: "__fortran_random_index".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                    ],
                    body: LambdaBody::Expr(Box::new(Expression::float(0.5))),
                    is_async: false,
                    captures: Vec::new(),
                }))],
                optional: false,
            })),
            else_: Box::new(Expression::float(0.5)),
        });
    }
    Expression::float(0.5)
}

fn walk_namelist_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let raw = pair.as_str().trim();
    let Some(rest) = raw.strip_prefix("namelist") else {
        return Ok(Statement::new(StmtKind::Block(vec![])));
    };

    let mut statements = Vec::new();
    let segments = rest
        .split('/')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>();
    for chunk in segments.chunks(2) {
        if chunk.len() != 2 {
            continue;
        }
        let group = chunk[0];
        let members = chunk[1]
            .split(',')
            .map(str::trim)
            .filter(|name| !name.is_empty())
            .collect::<Vec<_>>();
        let mut args = vec![Argument::positional(Expression::string(group))];
        args.extend(
            members
                .iter()
                .map(|member| Argument::positional(Expression::ident(member))),
        );
        statements.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__fortran_namelist_decl")),
                args,
                optional: false,
            },
        ))));
    }

    if statements.len() == 1 {
        Ok(statements
            .pop()
            .unwrap_or_else(|| Statement::new(StmtKind::Block(vec![]))))
    } else {
        Ok(Statement::new(StmtKind::Block(statements)))
    }
}

fn lower_fortran_namelist_io(body: &mut Vec<Statement>) {
    let mut groups = HashMap::new();
    lower_fortran_namelist_io_with_groups(body, &mut groups);
}

fn lower_fortran_namelist_io_with_groups(
    body: &mut Vec<Statement>,
    groups: &mut HashMap<String, Vec<String>>,
) {
    let mut lowered = Vec::with_capacity(body.len());
    for mut statement in body.drain(..) {
        match &mut statement.kind {
            StmtKind::Expr(expr) => {
                if let Some((group, members)) = parse_fortran_namelist_decl(expr) {
                    groups.insert(group, members);
                    continue;
                }
                if let Some(rewritten) = lower_fortran_namelist_helper(expr, groups) {
                    lowered.extend(rewritten);
                    continue;
                }
            }
            StmtKind::Block(stmts) => {
                lower_fortran_namelist_io_with_groups(stmts, groups);
            }
            StmtKind::FunctionDecl { body: nested, .. } => {
                let mut nested_groups = groups.clone();
                lower_fortran_namelist_io_with_groups(nested, &mut nested_groups);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_groups = groups.clone();
                lower_fortran_namelist_io_with_groups(then_body, &mut then_groups);
                for (_, elif_body) in elifs {
                    let mut elif_groups = groups.clone();
                    lower_fortran_namelist_io_with_groups(elif_body, &mut elif_groups);
                }
                if let Some(else_body) = else_body {
                    let mut else_groups = groups.clone();
                    lower_fortran_namelist_io_with_groups(else_body, &mut else_groups);
                }
            }
            _ => {}
        }
        lowered.push(statement);
    }
    *body = lowered;
}

fn parse_fortran_namelist_decl(expr: &Expression) -> Option<(String, Vec<String>)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !name.eq_ignore_ascii_case("__fortran_namelist_decl") {
        return None;
    }
    let Expression {
        kind: ExprKind::Lit(Literal::Str(group)),
        ..
    } = &args.first()?.value
    else {
        return None;
    };
    let members = args
        .iter()
        .skip(1)
        .filter_map(|arg| match &arg.value.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        })
        .collect::<Vec<_>>();
    Some((group.to_ascii_lowercase(), members))
}

fn lower_fortran_namelist_helper(
    expr: &Expression,
    groups: &HashMap<String, Vec<String>>,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };

    if name.eq_ignore_ascii_case("__fortran_namelist_write") {
        let file_number = args.first()?.value.clone();
        let Expression {
            kind: ExprKind::Lit(Literal::Str(group)),
            ..
        } = &args.get(1)?.value
        else {
            return None;
        };
        let members = groups.get(&group.to_ascii_lowercase())?;
        let mut statements = vec![Statement::new(StmtKind::PrintFile {
            file_number: file_number.clone(),
            items: vec![Expression::string(&format!("&{}", group))],
        })];
        for member in members {
            statements.push(Statement::new(StmtKind::PrintFile {
                file_number: file_number.clone(),
                items: vec![concat_fortran_io_parts(vec![
                    Expression::string(&format!(" {} = ", member)),
                    stringify_fortran_io_expr(Expression::ident(member)),
                    Expression::string(","),
                ])],
            }));
        }
        statements.push(Statement::new(StmtKind::PrintFile {
            file_number,
            items: vec![Expression::string("/")],
        }));
        return Some(statements);
    }

    if name.eq_ignore_ascii_case("__fortran_namelist_read") {
        let file_number = args.first()?.value.clone();
        let Expression {
            kind: ExprKind::Lit(Literal::Str(group)),
            ..
        } = &args.get(1)?.value
        else {
            return None;
        };
        let members = groups.get(&group.to_ascii_lowercase())?;
        let header_name = format!("__fortran_nml_header_{}", group);
        let footer_name = format!("__fortran_nml_footer_{}", group);
        let mut statements = vec![
            build_fortran_namelist_temp_decl(&header_name),
            Statement::new(StmtKind::LineInput {
                file_number: file_number.clone(),
                variable: header_name,
            }),
        ];
        for (index, member) in members.iter().enumerate() {
            let line_name = format!("__fortran_nml_line_{}_{}", group, index);
            statements.push(build_fortran_namelist_temp_decl(&line_name));
            statements.push(Statement::new(StmtKind::LineInput {
                file_number: file_number.clone(),
                variable: line_name.clone(),
            }));
            statements.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(member)],
                value: build_fortran_namelist_value_expr(&line_name),
                by_ref: false,
            }));
        }
        statements.push(build_fortran_namelist_temp_decl(&footer_name));
        statements.push(Statement::new(StmtKind::LineInput {
            file_number: file_number.clone(),
            variable: footer_name,
        }));
        if let Some(iostat_target) = args.get(2) {
            statements.push(Statement::new(StmtKind::Assign {
                targets: vec![iostat_target.value.clone()],
                value: Expression::int(0),
                by_ref: false,
            }));
        }
        return Some(statements);
    }

    None
}

fn build_fortran_namelist_temp_decl(name: &str) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: Some("character(len=4096)".to_string().into()),
            init: Some(Expression::string("")),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Dim,
    })
}

fn build_fortran_namelist_value_expr(line_name: &str) -> Expression {
    let split_eq = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("str_split")),
        args: vec![
            Argument::positional(Expression::ident(line_name)),
            Argument::positional(Expression::string("=")),
        ],
        optional: false,
    });
    let rhs = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("trim")),
        args: vec![Argument::positional(Expression::new(ExprKind::Index {
            object: Box::new(split_eq),
            index: Box::new(Expression::int(2)),
            null_safe: false,
        }))],
        optional: false,
    });
    let split_comma = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("str_split")),
        args: vec![
            Argument::positional(rhs),
            Argument::positional(Expression::string(",")),
        ],
        optional: false,
    });
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("trim")),
        args: vec![Argument::positional(Expression::new(ExprKind::Index {
            object: Box::new(split_comma),
            index: Box::new(Expression::int(1)),
            null_safe: false,
        }))],
        optional: false,
    })
}

fn is_fortran_allocator_intrinsic_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("allocate") || name.eq_ignore_ascii_case("deallocate"))
    )
}

fn lower_open_intrinsic_statement(expr: &Expression, args: &[Argument]) -> Option<Statement> {
    let file_number_index = args.iter().position(|arg| {
        arg.name.as_deref().is_some_and(|name| {
            name.eq_ignore_ascii_case("unit") || name.eq_ignore_ascii_case("newunit")
        })
    });
    let file_number_arg = file_number_index
        .and_then(|index| args.get(index))
        .or_else(|| args.first());
    let file_number = file_number_arg?.value.clone();
    // `open(61)` spells the unit POSITIONALLY, so there is no `unit=` to find
    // and `file_number_index` is None — which made the path search below,
    // which only skipped `file_number_index`, accept argument 0 and use the
    // UNIT NUMBER as the filename. That is where the stray files named `61`,
    // `57`, `32` … in the working directory came from.
    let unit_index = file_number_index.or(Some(0));

    let path = args
        .iter()
        .find(|arg| {
            arg.name
                .as_deref()
                .is_some_and(|name| name.eq_ignore_ascii_case("file"))
        })
        .map(|arg| arg.value.clone())
        .or_else(|| {
            args.iter()
                .enumerate()
                .find(|(index, arg)| Some(*index) != unit_index && arg.name.is_none())
                .map(|(_, arg)| arg.value.clone())
        })
        .unwrap_or_else(|| {
            if fortran_open_is_scratch(args) {
                // A SCRATCH file has no name by definition, and is deleted on
                // close — a per-site temp name is the right answer here.
                Expression::string(&format!(
                    "__fortran_scratch_{}_{}.tmp",
                    expr.span.start_line.max(1),
                    expr.span.start_col.max(1),
                ))
            } else {
                // An unnamed unit connects to `fort.<unit>` — gfortran, ifort
                // and every other compiler agree, and tests that write to a
                // bare unit and read it back depend on the two spellings
                // naming the same file. Built as a concatenation so a unit
                // held in a variable works too.
                // CONCAT, not Add: `Add` on a string and a number is numeric
                // addition, which named the file `NaN`.
                Expression::new(ExprKind::Binary {
                    op: BinOp::Concat,
                    left: Box::new(Expression::string("fort.")),
                    right: Box::new(file_number.clone()),
                })
            }
        });

    let mode = infer_fortran_open_mode(args);
    let open_stmt = Statement::new(StmtKind::OpenFile {
        path,
        mode,
        file_number: file_number.clone(),
    });

    if file_number_arg
        .and_then(|arg| arg.name.as_deref())
        .is_some_and(|name| name.eq_ignore_ascii_case("newunit"))
    {
        let assigned_unit = Expression::int(
            (expr.span.start_line.max(1) as i64) * 1000 + expr.span.start_col.max(1) as i64,
        );
        return Some(Statement::new(StmtKind::Block(vec![
            Statement::new(StmtKind::Assign {
                targets: vec![file_number],
                value: assigned_unit,
                by_ref: false,
            }),
            open_stmt,
        ])));
    }

    Some(open_stmt)
}

/// Whether this `open` asked for a SCRATCH file — one with no name, deleted
/// when it is closed.
fn fortran_open_is_scratch(args: &[Argument]) -> bool {
    args.iter().any(|arg| {
        arg.name
            .as_deref()
            .is_some_and(|name| name.eq_ignore_ascii_case("status"))
            && matches!(&arg.value.kind,
                ExprKind::Lit(Literal::Str(value)) if value.eq_ignore_ascii_case("scratch"))
    })
}

fn infer_fortran_open_mode(args: &[Argument]) -> FileMode {
    for arg in args {
        let Some(name) = arg.name.as_deref() else {
            continue;
        };
        let ExprKind::Lit(Literal::Str(value)) = &arg.value.kind else {
            continue;
        };
        let value = value.to_ascii_lowercase();
        if name.eq_ignore_ascii_case("position") && value == "append" {
            return FileMode::Append;
        }
        if name.eq_ignore_ascii_case("action") {
            return match value.as_str() {
                "read" => FileMode::Input,
                "write" | "readwrite" => FileMode::Output,
                _ => FileMode::Output,
            };
        }
        if name.eq_ignore_ascii_case("access") && value == "stream" {
            return FileMode::Binary;
        }
        if name.eq_ignore_ascii_case("form") && value == "unformatted" {
            return FileMode::Binary;
        }
        if name.eq_ignore_ascii_case("status") {
            return match value.as_str() {
                "old" => FileMode::Input,
                "scratch" | "replace" | "new" | "unknown" => FileMode::Output,
                _ => FileMode::Output,
            };
        }
    }
    FileMode::Output
}

fn bind_fortran_param_declarations(params: &mut [Param], body: &mut Vec<Statement>) {
    let mut rewritten = Vec::with_capacity(body.len());
    for mut statement in body.drain(..) {
        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                let mut remaining = Vec::new();
                for declaration in declarations.drain(..) {
                    let mut matched_param = false;
                    if let BindingPattern::Ident(name) = &declaration.pattern {
                        if let Some(param) = params
                            .iter_mut()
                            .find(|param| param.name.eq_ignore_ascii_case(name))
                        {
                            let declared_rank =
                                fortran_declared_array_rank(declaration.array_bounds.as_deref());
                            let declaration_type_hint =
                                declaration.type_hint.as_ref().map(|type_hint| {
                                    fortran_array_type_hint(
                                        type_hint,
                                        declaration.array_bounds.as_deref(),
                                    )
                                });
                            if param.type_hint.is_none()
                                || declared_rank
                                    > param
                                        .type_hint
                                        .as_deref()
                                        .map(fortran_type_hint_array_rank)
                                        .unwrap_or(0)
                            {
                                param.type_hint = declaration_type_hint.map(Into::into);
                            }
                            if matches!(param.pass_by, PassBy::Value)
                                && (declaration.array_bounds.is_some()
                                    || declaration
                                        .type_hint
                                        .as_deref()
                                        .and_then(parse_derived_type_name)
                                        .is_some())
                            {
                                param.pass_by = PassBy::Const;
                            }
                            matched_param = true;
                        }
                    }
                    if !matched_param {
                        remaining.push(declaration);
                    }
                }

                if !remaining.is_empty() {
                    *declarations = remaining;
                    rewritten.push(statement);
                }
            }
            _ => rewritten.push(statement),
        }
    }
    *body = rewritten;
}

fn apply_fortran_param_declaration_modes(params: &mut [Param], rest: &[Pair<Rule>]) {
    for pair in rest {
        apply_fortran_param_declaration_modes_from_pair(params, pair.clone());
    }
}

fn apply_fortran_param_declaration_modes_from_pair(params: &mut [Param], pair: Pair<Rule>) {
    if pair.as_rule() == Rule::var_declaration {
        let mut inner = pair.clone().into_inner();
        let type_hint = inner.next().map(|child| child.as_str().trim().to_string());
        let mut intent_mode = None;

        for child in inner {
            if child.as_rule() != Rule::var_attributes {
                continue;
            }

            for attr in child.into_inner().filter(|attr| meaningful(attr)) {
                if attr.as_rule() != Rule::var_attribute {
                    continue;
                }

                if let Some(mode) = parse_fortran_intent_mode(attr.as_str()) {
                    intent_mode = Some(mode);
                }
            }
        }

        if type_hint.is_some() || intent_mode.is_some() {
            for declaration in pair.clone().into_inner().filter(|child| meaningful(child)) {
                if declaration.as_rule() != Rule::var_declarator_list {
                    continue;
                }

                for declarator in declaration.into_inner().filter(|child| meaningful(child)) {
                    if declarator.as_rule() != Rule::var_declarator {
                        continue;
                    }

                    let mut parts = declarator.into_inner().filter(|child| meaningful(child));
                    let Some(name) = parts.next().map(|child| child.as_str().to_string()) else {
                        continue;
                    };

                    let Some(param) = params
                        .iter_mut()
                        .find(|param| param.name.eq_ignore_ascii_case(&name))
                    else {
                        continue;
                    };

                    if param.type_hint.is_none() {
                        param.type_hint = type_hint.clone().map(Into::into);
                    }

                    if let Some(mode) = intent_mode {
                        param.pass_by = mode;
                    }
                }
            }
        }
    }

    if pair.as_rule() == Rule::procedure_decl {
        let mut interface_name: Option<String> = None;
        let mut intent_mode = None;

        for child in pair.clone().into_inner().filter(|child| meaningful(child)) {
            match child.as_rule() {
                Rule::identifier if interface_name.is_none() => {
                    interface_name = Some(child.as_str().trim().to_string());
                }
                Rule::proc_attribute => {
                    if let Some(mode) = parse_fortran_intent_mode(child.as_str()) {
                        intent_mode = Some(mode);
                    }
                }
                _ => {}
            }
        }

        let type_hint = interface_name
            .as_ref()
            .map(|iface| format!("procedure({iface})"))
            .unwrap_or_else(|| "procedure".to_string());

        for declaration in pair.clone().into_inner().filter(|child| meaningful(child)) {
            if declaration.as_rule() != Rule::proc_decl_item {
                continue;
            }

            let Some(name) = declaration
                .into_inner()
                .filter(|child| meaningful(child))
                .find(|child| child.as_rule() == Rule::identifier)
                .map(|child| child.as_str().to_string())
            else {
                continue;
            };

            let Some(param) = params
                .iter_mut()
                .find(|param| param.name.eq_ignore_ascii_case(&name))
            else {
                continue;
            };

            if param.type_hint.is_none() {
                param.type_hint = Some(type_hint.clone().into());
            }

            if let Some(mode) = intent_mode {
                param.pass_by = mode;
            }
        }
    }

    for child in pair.into_inner().filter(|child| meaningful(child)) {
        apply_fortran_param_declaration_modes_from_pair(params, child);
    }
}

fn parse_fortran_intent_mode(attr_text: &str) -> Option<PassBy> {
    let attr = attr_text.trim().to_ascii_lowercase();
    if !attr.starts_with("intent(") {
        return None;
    }

    let inner = attr.strip_prefix("intent(")?.strip_suffix(')')?.trim();
    match inner {
        "in" => Some(PassBy::Const),
        "out" => Some(PassBy::Out),
        // Fortran passes dummy arguments by reference: the callee writes the
        // caller's storage directly, so the mutation is visible immediately
        // rather than being copied back on return. `Alias`, not `Ref` — the
        // same migration pascal `var`, vb `ByRef`, c# `ref` and php `&` took.
        "inout" => Some(PassBy::Alias),
        _ => None,
    }
}

/// Names a dummy declared with the VALUE attribute.
const FORTRAN_VALUE_DUMMY_MARKER: &str = "__fortran_value_dummy";

fn collect_fortran_value_dummies(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        if let StmtKind::Expr(expr) = &statement.kind {
            if let ExprKind::Call { callee, args, .. } = &expr.kind {
                if matches!(&callee.kind, ExprKind::Ident(n) if n == FORTRAN_VALUE_DUMMY_MARKER) {
                    for arg in args {
                        if let ExprKind::Lit(Literal::Str(name)) = &arg.value.kind {
                            out.insert(name.to_ascii_lowercase());
                        }
                    }
                }
            }
        }
        for_each_fortran_nested_body(&statement.kind, &mut |stmts| {
            collect_fortran_value_dummies(stmts, out)
        });
    }
}

/// Drop the VALUE-dummy markers once the promotion pass has read them.
///
/// The promotion runs per procedure DURING the walk, so the marker has to
/// survive until then; left in the tree afterwards it is a call to a name that
/// does not exist.
fn strip_fortran_value_dummy_markers(body: &mut Vec<Statement>) {
    for statement in body.iter_mut() {
        if let StmtKind::FunctionDecl { body, .. } = &mut statement.kind {
            strip_fortran_value_dummy_markers(body);
        }
        if let StmtKind::ModuleDecl { members, .. } | StmtKind::ClassDecl { members, .. } =
            &mut statement.kind
        {
            for member in members.iter_mut() {
                if let ClassMember::Method(inner) | ClassMember::NestedType(inner) = member {
                    strip_fortran_value_dummy_markers(&mut vec![(**inner).clone()]);
                    let mut one = vec![(**inner).clone()];
                    strip_fortran_value_dummy_markers(&mut one);
                    if let Some(first) = one.into_iter().next() {
                        **inner = first;
                    }
                }
            }
        }
        for_each_fortran_nested_vec_mut(&mut statement.kind, &mut strip_fortran_value_dummy_markers);
    }
    body.retain(|statement| {
        let StmtKind::Expr(expr) = &statement.kind else {
            return true;
        };
        let ExprKind::Call { callee, .. } = &expr.kind else {
            return true;
        };
        !matches!(&callee.kind, ExprKind::Ident(n) if n == FORTRAN_VALUE_DUMMY_MARKER)
    });
}

fn promote_mutated_fortran_params(params: &mut [Param], body: &[Statement]) {
    let mut by_value = HashSet::new();
    collect_fortran_value_dummies(body, &mut by_value);
    for param in params.iter_mut() {
        if by_value.contains(&param.name.to_ascii_lowercase()) {
            continue;
        }
        // A dummy argument with NO `intent` is by reference in Fortran, and it
        // lands here as `Value` — so the old `!= Const` guard skipped exactly
        // the case the language leaves implicit. MEASURED against gfortran:
        //
        //   subroutine bump(n)   ! no intent
        //     integer :: n
        //     n = n + 1
        //   end subroutine
        //
        // gfortran prints 2; `Value` printed 1, while `intent(inout)` and
        // `intent(out)` were both already correct.
        if !matches!(param.pass_by, PassBy::Const | PassBy::Value) {
            continue;
        }
        if body
            .iter()
            .any(|statement| statement_mutates_fortran_param(statement, &param.name))
        {
            param.pass_by = PassBy::Alias;
        }
    }
}

fn statement_mutates_fortran_param(statement: &Statement, param_name: &str) -> bool {
    match &statement.kind {
        StmtKind::Assign { targets, .. } => targets
            .iter()
            .any(|target| expr_targets_fortran_param(target, param_name)),
        // A named construct wraps its statement in `Labeled`; the wrapper is
        // transparent, so the question passes through to whatever it names.
        StmtKind::Labeled { body, .. } => statement_mutates_fortran_param(body, param_name),
        StmtKind::Block(stmts)
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::With { body: stmts, .. }
        | StmtKind::Using { body: stmts, .. }
        | StmtKind::Lock { body: stmts, .. } => stmts
            .iter()
            .any(|stmt| statement_mutates_fortran_param(stmt, param_name)),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            then_body
                .iter()
                .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || elifs.iter().any(|(_, body)| {
                    body.iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
        }
        StmtKind::While {
            body, else_body, ..
        }
        | StmtKind::ForIn {
            body, else_body, ..
        } => {
            body.iter()
                .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || else_body.as_ref().is_some_and(|stmts| {
                    stmts
                        .iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
        }
        StmtKind::For { init, body, .. } => {
            init.as_ref()
                .is_some_and(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || body
                    .iter()
                    .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
        }
        StmtKind::Switch { cases, default, .. } => {
            cases.iter().any(|case| {
                case.body
                    .iter()
                    .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
            }) || default.as_ref().is_some_and(|body| {
                body.iter()
                    .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
            })
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            body.iter()
                .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || catches.iter().any(|catch| {
                    catch
                        .body
                        .iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
                || else_body.as_ref().is_some_and(|stmts| {
                    stmts
                        .iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
                || finally.as_ref().is_some_and(|stmts| {
                    stmts
                        .iter()
                        .any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                })
        }
        _ => false,
    }
}

fn expr_targets_fortran_param(expr: &Expression, param_name: &str) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => name.eq_ignore_ascii_case(param_name),
        ExprKind::Index { object, .. } | ExprKind::Member { object, .. } => {
            expr_targets_fortran_param(object, param_name)
        }
        _ => false,
    }
}

fn normalize_fortran_function_result(
    function_name: &str,
    result_name: Option<&str>,
    return_type: &mut Option<String>,
    body: &mut Vec<Statement>,
) {
    let result_var = result_name.unwrap_or(function_name).to_string();

    if return_type.is_none() {
        *return_type = find_function_result_type(body, &result_var);
    }

    rewrite_function_returns(body, &result_var);

    // Ensure the result variable is declared as a local — required when the function
    // name itself is used as the result slot (e.g. `integer function cube(n)` where
    // the body does `cube = n*n*n`).
    let result_var_declared = body.iter().any(|stmt| {
        if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
            declarations.iter().any(|d| {
                matches!(&d.pattern, BindingPattern::Ident(n) if n.eq_ignore_ascii_case(&result_var))
            })
        } else {
            false
        }
    });
    if !result_var_declared {
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(result_var.clone()),
                    type_hint: return_type.clone().map(Into::into),
                    init: None,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
        );
    }

    let needs_final_return = !matches!(
        body.last().map(|stmt| &stmt.kind),
        Some(StmtKind::Return(_))
    );
    if needs_final_return {
        body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
            &result_var,
        )))));
    }
}

fn find_function_result_type(body: &[Statement], result_var: &str) -> Option<String> {
    body.iter().find_map(|statement| {
        let StmtKind::VarDecl { declarations, .. } = &statement.kind else {
            return None;
        };
        declarations.iter().find_map(|declaration| {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                return None;
            };
            if name.eq_ignore_ascii_case(result_var) {
                declaration.type_hint.as_ref().map(|type_hint| {
                    fortran_array_type_hint(type_hint, declaration.array_bounds.as_deref())
                })
            } else {
                None
            }
        })
    })
}

fn rewrite_function_returns(body: &mut [Statement], result_var: &str) {
    for statement in body.iter_mut() {
        match &mut statement.kind {
            StmtKind::Return(value)
                if value
                    .as_ref()
                    .is_none_or(|expr| matches!(expr.kind, ExprKind::Lit(Literal::Null))) =>
            {
                *value = Some(Expression::ident(result_var));
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => rewrite_function_returns(stmts, result_var),
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                rewrite_function_returns(std::slice::from_mut(body.as_mut()), result_var)
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_function_returns(then_body, result_var);
                for (_, elif_body) in elifs {
                    rewrite_function_returns(elif_body, result_var);
                }
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                rewrite_function_returns(stmts, result_var);
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                if let Some(init) = init {
                    rewrite_function_returns(std::slice::from_mut(init.as_mut()), result_var);
                }
                rewrite_function_returns(stmts, result_var);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    rewrite_function_returns(&mut case.body, result_var);
                }
                if let Some(default) = default {
                    rewrite_function_returns(default, result_var);
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_function_returns(try_body, result_var);
                for catch in catches {
                    rewrite_function_returns(&mut catch.body, result_var);
                }
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
                if let Some(finally) = finally {
                    rewrite_function_returns(finally, result_var);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_semantics(params: &[Param], body: &mut Vec<Statement>) {
    let mut arrays = HashSet::new();
    let mut char_vars = HashSet::new();
    let mut callables = HashSet::new();
    let mut array_fields = HashSet::new();
    for param in params {
        if param
            .type_hint
            .as_deref()
            .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
        {
            arrays.insert(param.name.to_ascii_lowercase());
        }
        if param
            .type_hint
            .as_deref()
            .is_some_and(is_fortran_string_type_hint)
        {
            char_vars.insert(param.name.to_ascii_lowercase());
            arrays.insert(param.name.to_ascii_lowercase());
        }
        if param
            .type_hint
            .as_deref()
            .is_some_and(is_fortran_callable_type_hint)
        {
            callables.insert(param.name.to_ascii_lowercase());
        }
    }
    collect_fortran_callable_names(body, &mut callables);
    collect_fortran_array_field_names(body, &mut array_fields);
    lower_fortran_array_semantics_with_env(
        body,
        &mut arrays,
        &mut char_vars,
        &mut callables,
        &array_fields,
    );
}

fn repair_remaining_fortran_array_calls(body: &mut [Statement]) {
    let mut arrays = HashSet::new();
    let mut callables = HashSet::new();
    let mut array_fields = HashSet::new();
    collect_fortran_callable_names(body, &mut callables);
    collect_fortran_array_field_names(body, &mut array_fields);
    repair_remaining_fortran_array_calls_with_env(body, &mut arrays, &mut callables, &array_fields);
}

fn repair_remaining_fortran_array_calls_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    callables: &mut HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_remaining_fortran_array_calls_in_statement(
            statement,
            arrays,
            callables,
            array_fields,
        );

        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration.array_bounds.is_some()
                    || declaration
                        .init
                        .as_ref()
                        .is_some_and(is_array_initializer_expr)
                    || declaration
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_string_type_hint)
                {
                    arrays.insert(name.to_ascii_lowercase());
                }
                if declaration
                    .type_hint
                    .as_deref()
                    .is_some_and(is_fortran_callable_type_hint)
                {
                    callables.insert(name.to_ascii_lowercase());
                }
            }
        }

        match &mut statement.kind {
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested = arrays.clone();
                let mut nested_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(
                    stmts,
                    &mut nested,
                    &mut nested_callables,
                    array_fields,
                );
            }
            // A NAMED construct — `outer: do …` / `outer: block …` — wraps its
            // statement in `Labeled`, and without an arm here the whole body
            // fell to `_ => {}`. Every `w(i)` inside a named loop kept its
            // `Call` shape and read the array 0-based, so naming a loop
            // silently broke subscripting in its entire body.
            //
            // The wrapper is TRANSPARENT: it introduces no scope of its own, so
            // the env passes straight through and the loop or block underneath
            // makes whatever scope it needs.
            StmtKind::Labeled { body, .. } => {
                repair_remaining_fortran_array_calls_with_env(
                    std::slice::from_mut(body.as_mut()),
                    arrays,
                    callables,
                    array_fields,
                );
            }
            StmtKind::ModuleDecl { members, .. } => {
                repair_remaining_fortran_array_calls_in_members(
                    members,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = arrays.clone();
                let mut then_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(
                    then_body,
                    &mut then_arrays,
                    &mut then_callables,
                    array_fields,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    let mut elif_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        elif_body,
                        &mut elif_arrays,
                        &mut elif_callables,
                        array_fields,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_callables,
                    array_fields,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                if let Some(init) = init {
                    repair_remaining_fortran_array_calls_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        &mut loop_callables,
                        array_fields,
                    );
                }
                repair_remaining_fortran_array_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_callables,
                    array_fields,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    let mut case_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        &mut case_callables,
                        array_fields,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    let mut default_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        default,
                        &mut default_arrays,
                        &mut default_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = arrays.clone();
                let mut try_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(
                    try_body,
                    &mut try_arrays,
                    &mut try_callables,
                    array_fields,
                );
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    let mut catch_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        &mut catch_callables,
                        array_fields,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_callables,
                        array_fields,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    let mut finally_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(
                        finally,
                        &mut finally_arrays,
                        &mut finally_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                let mut fn_arrays = arrays.clone();
                let mut fn_callables = callables.clone();
                for param in params.iter() {
                    if param
                        .type_hint
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        fn_arrays.insert(param.name.to_ascii_lowercase());
                    }
                    if param
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_callable_type_hint)
                    {
                        fn_callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                collect_fortran_callable_names(body, &mut fn_callables);
                repair_remaining_fortran_array_calls_with_env(
                    body,
                    &mut fn_arrays,
                    &mut fn_callables,
                    array_fields,
                );
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                repair_remaining_fortran_array_calls_in_members(
                    members,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            _ => {}
        }
    }
}

fn repair_remaining_fortran_array_calls_in_members(
    members: &mut [ClassMember],
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { params, body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut method_arrays = arrays.clone();
                let mut method_callables = callables.clone();
                for param in params.iter() {
                    if param
                        .type_hint
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        method_arrays.insert(param.name.to_ascii_lowercase());
                    }
                    if param
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_callable_type_hint)
                    {
                        method_callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                repair_remaining_fortran_array_calls_with_env(
                    body,
                    &mut method_arrays,
                    &mut method_callables,
                    array_fields,
                );
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = arrays.clone();
                let mut nested_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested_arrays,
                    &mut nested_callables,
                    array_fields,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_remaining_fortran_array_calls_in_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            if !is_fortran_allocator_intrinsic_expr(expr) {
                rewrite_remaining_fortran_array_calls_in_expr(
                    expr,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_remaining_fortran_array_calls_in_expr(
                    target,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(target, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::Return(Some(expr)) => {
            rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields)
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_remaining_fortran_array_calls_in_expr(
                    expr,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            if let Some(cause) = cause {
                rewrite_remaining_fortran_array_calls_in_expr(
                    cause,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        StmtKind::If { cond, elifs, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
            // The elif CONDITIONS, which the `_with_env` arm above discards
            // (`for (_, elif_body)`) because it only recurses into bodies.
            // Nothing else repairs them, so `else if (x /= w(i))` kept its
            // `Call` even when the plain `if` was fixed.
            for (elif_cond, _) in elifs.iter_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    elif_cond,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. }
        | StmtKind::Switch { expr: cond, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_remaining_fortran_array_calls_in_expr(
                    cond,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            if let Some(update) = update {
                rewrite_remaining_fortran_array_calls_in_expr(
                    update,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        _ => {}
    }
}

fn rewrite_remaining_fortran_array_calls_in_expr(
    expr: &mut Expression,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(left, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(right, arrays, callables, array_fields);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner) => {
            rewrite_remaining_fortran_array_calls_in_expr(inner, arrays, callables, array_fields)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(then, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(else_, arrays, callables, array_fields);
        }
        ExprKind::Member { object, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(object, arrays, callables, array_fields)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(object, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(index, arrays, callables, array_fields);
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    lower,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    upper,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            if let Some(step) = step.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    step,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            rewrite_remaining_fortran_array_calls_in_expr(callee, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    &mut arg.value,
                    arrays,
                    callables,
                    array_fields,
                );
            }
            if !args.is_empty()
                && !*optional
                && !is_known_fortran_callable(callee, callables)
                && (is_known_fortran_array(callee, arrays, array_fields)
                    || matches!(&callee.kind, ExprKind::Index { .. }))
            {
                expr.kind = build_fortran_index_chain(callee.as_ref().clone(), args);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_remaining_fortran_array_calls_in_expr(class, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(
                    &mut arg.value,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_remaining_fortran_array_calls_in_expr(target, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        ExprKind::Object(items) => {
            for item in items {
                match item {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_remaining_fortran_array_calls_in_expr(
                            key,
                            arrays,
                            callables,
                            array_fields,
                        );
                        rewrite_remaining_fortran_array_calls_in_expr(
                            value,
                            arrays,
                            callables,
                            array_fields,
                        );
                    }
                    ObjectProperty::Spread(expr) => rewrite_remaining_fortran_array_calls_in_expr(
                        expr,
                        arrays,
                        callables,
                        array_fields,
                    ),
                    ObjectProperty::Shorthand(_)
                    | ObjectProperty::Method { .. }
                    | ObjectProperty::Accessor { .. } => {}
                }
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = item.key.as_mut() {
                    rewrite_remaining_fortran_array_calls_in_expr(
                        key,
                        arrays,
                        callables,
                        array_fields,
                    );
                }
                rewrite_remaining_fortran_array_calls_in_expr(
                    &mut item.value,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args {
                rewrite_remaining_fortran_array_calls_in_expr(
                    arg,
                    arrays,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    rewrite_remaining_fortran_array_calls_in_expr(
                        expr,
                        arrays,
                        callables,
                        array_fields,
                    );
                }
            }
        }
        _ => {}
    }
}

fn collect_fortran_array_field_names(body: &[Statement], array_fields: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_names_in_members(members, array_fields);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_field_names_in_members(members, array_fields)
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_fortran_array_field_names(body, array_fields)
            }
            _ => {}
        }
    }
}

fn collect_fortran_callable_names(body: &[Statement], callables: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    if declaration
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_callable_type_hint)
                    {
                        callables.insert(name.to_ascii_lowercase());
                    }
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                for param in params {
                    if param
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_callable_type_hint)
                    {
                        callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                collect_fortran_callable_names(body, callables);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            collect_fortran_callable_names(
                                std::slice::from_ref(stmt.as_ref()),
                                callables,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_callable_names(body, callables),
            _ => {}
        }
    }
}

fn collect_fortran_array_function_names(body: &[Statement], array_functions: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::FunctionDecl {
                name,
                return_type,
                body,
                ..
            } => {
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    array_functions.insert(name.to_ascii_lowercase());
                }
                collect_fortran_array_function_names(body, array_functions);
            }
            StmtKind::InterfaceDecl { members, .. } => {
                for member in members {
                    if let InterfaceMember::Method {
                        name,
                        return_type,
                        signature_source,
                        ..
                    } = member
                    {
                        if return_type
                            .as_deref()
                            .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        {
                            array_functions.insert(name.to_ascii_lowercase());
                        }
                        if signature_source.as_ref().is_some_and(|source| {
                            array_functions.contains(&source.to_ascii_lowercase())
                        }) {
                            array_functions.insert(name.to_ascii_lowercase());
                        }
                    }
                }
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_function_names_in_members(members, array_functions);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_function_names_in_members(members, array_functions);
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_fortran_array_function_names(body, array_functions)
            }
            _ => {}
        }
    }
}

fn collect_fortran_elemental_function_names(
    body: &[Statement],
    elemental_functions: &mut HashSet<String>,
) {
    for statement in body {
        match &statement.kind {
            StmtKind::FunctionDecl {
                name,
                modifiers,
                body,
                ..
            } => {
                if modifiers.decorators.iter().any(|decorator| {
                    matches!(&decorator.kind, ExprKind::Ident(value) if value.eq_ignore_ascii_case("elemental"))
                }) {
                    elemental_functions.insert(name.to_ascii_lowercase());
                }
                collect_fortran_elemental_function_names(body, elemental_functions);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::Method(stmt) | ClassMember::NestedType(stmt) = member {
                        collect_fortran_elemental_function_names(
                            std::slice::from_ref(stmt.as_ref()),
                            elemental_functions,
                        );
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_fortran_elemental_function_names(body, elemental_functions);
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_function_names_in_members(
    members: &[ClassMember],
    array_functions: &mut HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                collect_fortran_array_function_names(
                    std::slice::from_ref(stmt.as_ref()),
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_sizes(
    body: &[Statement],
    array_field_sizes: &mut HashMap<String, Expression>,
) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_sizes_in_members(members, array_field_sizes);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_field_sizes_in_members(members, array_field_sizes);
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_fortran_array_field_sizes(body, array_field_sizes)
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_sizes_in_members(
    members: &[ClassMember],
    array_field_sizes: &mut HashMap<String, Expression>,
) {
    for member in members {
        match member {
            ClassMember::Field {
                name,
                array_bounds,
                init,
                ..
            } => {
                if let Some(size) = array_bounds
                    .as_deref()
                    .and_then(bounds_total_size_expr)
                    .or_else(|| init.as_ref().and_then(array_init_size_expr))
                {
                    array_field_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_sizes(
                    std::slice::from_ref(stmt.as_ref()),
                    array_field_sizes,
                );
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_names_in_members(
    members: &[ClassMember],
    array_fields: &mut HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Field {
                name,
                type_hint,
                array_bounds,
                ..
            } => {
                if array_bounds.is_some()
                    || type_hint
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                    || type_hint
                        .as_deref()
                        .is_some_and(is_fortran_string_type_hint)
                {
                    array_fields.insert(name.to_ascii_lowercase());
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_names(
                    std::slice::from_ref(stmt.as_ref()),
                    array_fields,
                );
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_semantics_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    char_vars: &mut HashSet<String>,
    callables: &mut HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        // Character substring assignment `s(l:r) = value` must be rewritten to string
        // concatenation BEFORE the generic subscript normaliser runs — otherwise the
        // normaliser treats `s` as an array and the downstream loop-lowering pass
        // produces wrong code for immutable JS strings.
        rewrite_fortran_char_slice_assign(statement, char_vars);

        rewrite_array_subscripts_in_statement(
            statement,
            arrays,
            char_vars,
            callables,
            array_fields,
        );

        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                let lower = name.to_ascii_lowercase();
                if declaration
                    .type_hint
                    .as_deref()
                    .is_some_and(is_fortran_string_type_hint)
                {
                    char_vars.insert(lower.clone());
                }
                if declaration.array_bounds.is_some()
                    || declaration
                        .init
                        .as_ref()
                        .is_some_and(is_array_initializer_expr)
                    || declaration
                        .type_hint
                        .as_deref()
                        .is_some_and(is_fortran_string_type_hint)
                {
                    arrays.insert(lower.clone());
                }
                if declaration
                    .type_hint
                    .as_deref()
                    .is_some_and(is_fortran_callable_type_hint)
                {
                    callables.insert(lower);
                }
            }
        }

        match &mut statement.kind {
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = arrays.clone();
                let mut nested_chars = char_vars.clone();
                let mut nested_callables = callables.clone();
                lower_fortran_array_semantics_with_env(
                    stmts,
                    &mut nested,
                    &mut nested_chars,
                    &mut nested_callables,
                    array_fields,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_semantics_with_env(
                    std::slice::from_mut(body.as_mut()),
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = arrays.clone();
                let mut then_chars = char_vars.clone();
                let mut then_callables = callables.clone();
                lower_fortran_array_semantics_with_env(
                    then_body,
                    &mut then_arrays,
                    &mut then_chars,
                    &mut then_callables,
                    array_fields,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    let mut elif_chars = char_vars.clone();
                    let mut elif_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        elif_body,
                        &mut elif_arrays,
                        &mut elif_chars,
                        &mut elif_callables,
                        array_fields,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_chars = char_vars.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_chars,
                        &mut else_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_chars = char_vars.clone();
                let mut loop_callables = callables.clone();
                lower_fortran_array_semantics_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_chars,
                    &mut loop_callables,
                    array_fields,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_chars = char_vars.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_chars,
                        &mut else_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_chars = char_vars.clone();
                let mut loop_callables = callables.clone();
                if let Some(init) = init {
                    lower_fortran_array_semantics_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        &mut loop_chars,
                        &mut loop_callables,
                        array_fields,
                    );
                }
                lower_fortran_array_semantics_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_chars,
                    &mut loop_callables,
                    array_fields,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    let mut case_chars = char_vars.clone();
                    let mut case_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        &mut case_chars,
                        &mut case_callables,
                        array_fields,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    let mut default_chars = char_vars.clone();
                    let mut default_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        default,
                        &mut default_arrays,
                        &mut default_chars,
                        &mut default_callables,
                        array_fields,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = arrays.clone();
                let mut try_chars = char_vars.clone();
                let mut try_callables = callables.clone();
                lower_fortran_array_semantics_with_env(
                    try_body,
                    &mut try_arrays,
                    &mut try_chars,
                    &mut try_callables,
                    array_fields,
                );
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    let mut catch_chars = char_vars.clone();
                    let mut catch_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        &mut catch_chars,
                        &mut catch_callables,
                        array_fields,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_chars = char_vars.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_chars,
                        &mut else_callables,
                        array_fields,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    let mut finally_chars = char_vars.clone();
                    let mut finally_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(
                        finally,
                        &mut finally_arrays,
                        &mut finally_chars,
                        &mut finally_callables,
                        array_fields,
                    );
                }
            }
            _ => {}
        }
    }
}

/// Rewrite `s(l:r) = value` into `s = s.slice(0, l-1) + value + s.slice(r)` when `s` is a
/// known character variable.  This must run before the generic subscript normaliser so that
/// the character slice target is removed before the array-assignment lowering sees it.
///
/// In `walk_assign`, LHS subscripts are built directly as Index nodes, so the target has the
/// form `Index { object: Ident("s"), index: Slice { lower, upper } }`.
fn build_fortran_str_slice(
    object: Expression,
    start: Expression,
    end: Option<Expression>,
) -> Expression {
    let end = end.unwrap_or_else(|| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("len")),
            args: vec![Argument::positional(object.clone())],
            optional: false,
        })
    });

    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(Expression::new(ExprKind::Range {
            start: Box::new(start),
            end: Box::new(end),
            inclusive: false,
        })),
        null_safe: false,
    })
}

fn rewrite_fortran_char_slice_assign(statement: &mut Statement, char_vars: &HashSet<String>) {
    let StmtKind::Assign { targets, value, .. } = &statement.kind else {
        return;
    };
    let [target] = targets.as_slice() else {
        return;
    };

    // `s(l:r)` on the LHS is Index { object: Ident("s"), index: Slice { lower, upper } }
    let ExprKind::Index { object, index, .. } = &target.kind else {
        return;
    };
    let ExprKind::Slice { lower, upper, .. } = &index.kind else {
        return;
    };
    let ExprKind::Ident(var_name) = &object.kind else {
        return;
    };
    if !char_vars.contains(&var_name.to_ascii_lowercase()) {
        return;
    }

    // Fortran indices are 1-based.  For `s(l:r) = v` the JS equivalent is:
    //   s = s.slice(0, l-1) + v + s.slice(r)
    // The slice end in JS is exclusive and already equals the Fortran 1-based upper bound.
    let var = Expression::ident(var_name);
    let val = value.clone();

    // prefix: characters before the replaced range
    let pre = match lower.as_deref() {
        None => Expression::string(""),
        Some(l) => build_fortran_str_slice(
            var.clone(),
            Expression::int(0),
            Some(Expression::new(ExprKind::Binary {
                left: Box::new(l.clone()),
                op: BinOp::Sub,
                right: Box::new(Expression::int(1)),
            })),
        ),
    };

    // suffix: characters after the replaced range
    let post = match upper.as_deref() {
        None => Expression::string(""),
        Some(r) => build_fortran_str_slice(var.clone(), r.clone(), None),
    };

    let new_value = Expression::new(ExprKind::Binary {
        left: Box::new(Expression::new(ExprKind::Binary {
            left: Box::new(pre),
            op: BinOp::Concat,
            right: Box::new(val),
        })),
        op: BinOp::Concat,
        right: Box::new(post),
    });

    let var_name = var_name.clone();
    *statement = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: new_value,
        by_ref: false,
    });
}

fn lower_fortran_array_assignments(body: &mut [Statement]) {
    let mut arrays = HashSet::new();
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    let mut elemental_functions = HashSet::new();
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    collect_fortran_elemental_function_names(body, &mut elemental_functions);
    lower_fortran_array_assignments_with_env(
        body,
        &mut arrays,
        &mut array_sizes,
        &array_field_sizes,
        &array_fields,
        &array_functions,
        &elemental_functions,
    );
}

fn lower_fortran_scalar_array_assignments(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_ranks = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_field_ranks = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();

    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_scalar_array_assignments_with_env(
        body,
        &mut array_sizes,
        &mut array_ranks,
        &array_fields,
        &array_field_sizes,
        &mut array_field_ranks,
        &array_functions,
    );
}

fn lower_fortran_array_call_arguments(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_call_arguments_with_env(
        body,
        &mut array_sizes,
        &array_field_sizes,
        &array_fields,
        &array_functions,
    );
}

fn lower_fortran_array_call_arguments_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_call_argument_statement(
            statement,
            array_sizes,
            array_field_sizes,
            array_fields,
            array_functions,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::FunctionDecl {
                body: function_body,
                ..
            } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    function_body,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_call_arguments_in_members(
                    members,
                    array_sizes,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    stmts,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_call_arguments_with_env(
                    std::slice::from_mut(body.as_mut()),
                    array_sizes,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    then_body,
                    &mut then_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        elif_body,
                        &mut elif_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_call_arguments_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
                lower_fortran_array_call_arguments_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        default,
                        &mut default_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    try_body,
                    &mut try_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(
                        finally,
                        &mut finally_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_call_arguments_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    body,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            ClassMember::NestedType(stmt) => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_call_argument_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    let (callee, args, optional, rebuild) = match &statement.kind {
        StmtKind::Expr(expr) => {
            if is_fortran_allocator_intrinsic_expr(expr) {
                return;
            }
            let ExprKind::Call {
                callee,
                args,
                optional,
            } = &expr.kind
            else {
                return;
            };
            (callee, args, optional, None)
        }
        StmtKind::Assign { targets, value, .. } => {
            let ExprKind::Call {
                callee,
                args,
                optional,
            } = &value.kind
            else {
                return;
            };
            (callee, args, optional, Some(targets.clone()))
        }
        _ => return,
    };
    if *optional {
        return;
    }

    let mut setup = Vec::new();
    let mut lowered_args = args.clone();
    let mut temp_index = 0usize;

    for arg in lowered_args.iter_mut() {
        if !expr_is_known_fortran_array(&arg.value, array_sizes, array_field_sizes, array_functions)
            || matches!(
                arg.value.kind,
                ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Array(_)
            )
        {
            continue;
        }
        let Some(size) =
            resolve_fortran_array_expr_size(&arg.value, array_sizes, array_field_sizes)
        else {
            continue;
        };
        let temp_name = format!("__fortran_array_arg_{temp_index}");
        temp_index += 1;
        setup.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(temp_name.clone()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![
                        Argument::positional(size.clone()),
                        Argument::positional(Expression::int(0)),
                    ],
                    optional: false,
                })),
                array_bounds: Some(vec![size.clone()]),
                with_events: false,
            }],
            kind: VarDeclKind::Dim,
        }));
        setup.push(build_fortran_array_materialization_statement(
            &temp_name,
            size,
            &arg.value,
            array_sizes,
            array_fields,
            array_functions,
        ));
        arg.value = Expression::ident(&temp_name);
    }

    if setup.is_empty() {
        return;
    }

    let lowered_call = Expression::new(ExprKind::Call {
        callee: callee.clone(),
        args: lowered_args,
        optional: false,
    });
    setup.push(match rebuild {
        Some(targets) => Statement::new(StmtKind::Assign {
            targets,
            value: lowered_call,
            by_ref: false,
        }),
        None => Statement::new(StmtKind::Expr(lowered_call)),
    });
    *statement = Statement::new(StmtKind::Block(setup));
}

fn resolve_fortran_array_expr_size(
    expr: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    if let Some(key) = fortran_array_target_key(expr) {
        if let Some(size) = array_sizes.get(&key) {
            return Some(size.clone());
        }
    }

    match &expr.kind {
        ExprKind::Ident(name) => array_sizes.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Member { field, .. } => {
            array_field_sizes.get(&field.to_ascii_lowercase()).cloned()
        }
        ExprKind::Array(items) => Some(Expression::int(items.len() as i64)),
        ExprKind::Binary { left, right, .. } => {
            resolve_fortran_array_expr_size(left, array_sizes, array_field_sizes)
                .or_else(|| resolve_fortran_array_expr_size(right, array_sizes, array_field_sizes))
        }
        ExprKind::Unary { expr: inner, .. } => {
            resolve_fortran_array_expr_size(inner, array_sizes, array_field_sizes)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            resolve_fortran_array_expr_size(cond, array_sizes, array_field_sizes)
                .or_else(|| resolve_fortran_array_expr_size(then, array_sizes, array_field_sizes))
                .or_else(|| resolve_fortran_array_expr_size(else_, array_sizes, array_field_sizes))
        }
        ExprKind::Slice { lower, upper, .. } => {
            let lower = lower
                .as_deref()
                .cloned()
                .unwrap_or_else(|| Expression::int(0));
            upper.as_deref().cloned().map(|upper| {
                Expression::new(ExprKind::Binary {
                    left: Box::new(upper),
                    op: BinOp::Sub,
                    right: Box::new(lower),
                })
            })
        }
        ExprKind::Index { object, index, .. } => match &index.kind {
            ExprKind::Slice { .. } => fortran_slice_extent(expr),
            _ => resolve_fortran_array_expr_size(object, array_sizes, array_field_sizes),
        },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Member { object, field, .. }
                if matches!(
                    field.to_ascii_lowercase().as_str(),
                    "map" | "filter" | "flatmap"
                ) =>
            {
                resolve_fortran_array_expr_size(object, array_sizes, array_field_sizes)
            }
            _ => None,
        },
        _ => None,
    }
}

fn build_fortran_array_materialization_statement(
    temp_name: &str,
    size: Expression,
    value: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Statement {
    let loop_var = "__fortran_array_arg_index";
    let loop_expr = Expression::ident(loop_var);
    Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(loop_var)],
            value: Expression::int(0),
            by_ref: false,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::ident(loop_var)),
            op: BinOp::Lt,
            right: Box::new(size),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(loop_var)),
            value: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::ident(loop_var)),
                op: BinOp::Add,
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: vec![Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(temp_name)),
                index: Box::new(loop_expr.clone()),
                null_safe: false,
            })],
            value: lower_fortran_array_materialization_value(
                value,
                &loop_expr,
                array_sizes,
                array_fields,
                array_functions,
            ),
            by_ref: false,
        })],
    })
}

fn lower_fortran_array_materialization_value(
    expr: &Expression,
    loop_index: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if array_sizes.contains_key(&name.to_ascii_lowercase()) => {
            Expression::new(ExprKind::Index {
                object: Box::new(expr.clone()),
                index: Box::new(loop_index.clone()),
                null_safe: false,
            })
        }
        ExprKind::Member { field, .. } if array_fields.contains(&field.to_ascii_lowercase()) => {
            Expression::new(ExprKind::Index {
                object: Box::new(expr.clone()),
                index: Box::new(loop_index.clone()),
                null_safe: false,
            })
        }
        ExprKind::Array(_) | ExprKind::Slice { .. } | ExprKind::ArrayTransform { .. } => {
            Expression::new(ExprKind::Index {
                object: Box::new(expr.clone()),
                index: Box::new(loop_index.clone()),
                null_safe: false,
            })
        }
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array") || array_functions.contains(&name.to_ascii_lowercase()))
                || matches!(&callee.kind, ExprKind::Member { field, .. } if matches!(field.to_ascii_lowercase().as_str(), "map" | "filter" | "flatmap")) =>
        {
            Expression::new(ExprKind::Index {
                object: Box::new(expr.clone()),
                index: Box::new(loop_index.clone()),
                null_safe: false,
            })
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(lower_fortran_array_materialization_value(
                object,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            index: Box::new(lower_fortran_array_materialization_value(
                index,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            null_safe: *null_safe,
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(lower_fortran_array_materialization_value(
                left,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            right: Box::new(lower_fortran_array_materialization_value(
                right,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(lower_fortran_array_materialization_value(
                inner,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(lower_fortran_array_materialization_value(
                cond,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            then: Box::new(lower_fortran_array_materialization_value(
                then,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            else_: Box::new(lower_fortran_array_materialization_value(
                else_,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        _ => expr.clone(),
    }
}

fn lower_fortran_array_assignments_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
    elemental_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_assignment_statement(
            statement,
            arrays,
            array_sizes,
            array_field_sizes,
            array_fields,
            array_functions,
            elemental_functions,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    if declaration.array_bounds.is_some()
                        || declaration
                            .init
                            .as_ref()
                            .is_some_and(is_array_initializer_expr)
                        || declaration
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        arrays.insert(name.to_ascii_lowercase());
                    }
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::Expr(expr) => {
                let mut ignored_ranks = HashMap::new();
                record_fortran_allocate_array_metadata(expr, array_sizes, &mut ignored_ranks, None);
            }
            StmtKind::FunctionDecl {
                params,
                body: function_body,
                ..
            } => {
                let mut nested_arrays = arrays.clone();
                for param in params {
                    let array_rank = param
                        .type_hint
                        .as_deref()
                        .map(fortran_type_hint_array_rank)
                        .unwrap_or(0);
                    if array_rank > 0
                        || param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        nested_arrays.insert(param.name.to_ascii_lowercase());
                    }
                }
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    function_body,
                    &mut nested_arrays,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_assignments_in_members(
                    members,
                    arrays,
                    array_sizes,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested_arrays = arrays.clone();
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    stmts,
                    &mut nested_arrays,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_assignments_with_env(
                    std::slice::from_mut(body.as_mut()),
                    arrays,
                    array_sizes,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_known_arrays = arrays.clone();
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    then_body,
                    &mut then_known_arrays,
                    &mut then_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
                for (_, elif_body) in elifs {
                    let mut elif_known_arrays = arrays.clone();
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        elif_body,
                        &mut elif_known_arrays,
                        &mut elif_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_known_arrays = arrays.clone();
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        else_body,
                        &mut else_known_arrays,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_known_arrays = arrays.clone();
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    stmts,
                    &mut loop_known_arrays,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
                if let Some(else_body) = else_body {
                    let mut else_known_arrays = arrays.clone();
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        else_body,
                        &mut else_known_arrays,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_known_arrays = arrays.clone();
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_assignments_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_known_arrays,
                        &mut loop_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
                lower_fortran_array_assignments_with_env(
                    stmts,
                    &mut loop_known_arrays,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_known_arrays = arrays.clone();
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        &mut case.body,
                        &mut case_known_arrays,
                        &mut case_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
                if let Some(default) = default {
                    let mut default_known_arrays = arrays.clone();
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        default,
                        &mut default_known_arrays,
                        &mut default_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_known_arrays = arrays.clone();
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    try_body,
                    &mut try_known_arrays,
                    &mut try_arrays,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
                for catch in catches {
                    let mut catch_known_arrays = arrays.clone();
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        &mut catch.body,
                        &mut catch_known_arrays,
                        &mut catch_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_known_arrays = arrays.clone();
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        else_body,
                        &mut else_known_arrays,
                        &mut else_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_known_arrays = arrays.clone();
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(
                        finally,
                        &mut finally_known_arrays,
                        &mut finally_arrays,
                        array_field_sizes,
                        array_fields,
                        array_functions,
                        elemental_functions,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_assignments_in_members(
    members: &mut [ClassMember],
    arrays: &HashSet<String>,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
    elemental_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut nested_arrays = arrays.clone();
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    body,
                    &mut nested_arrays,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = arrays.clone();
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested_arrays,
                    &mut nested,
                    array_field_sizes,
                    array_fields,
                    array_functions,
                    elemental_functions,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_assignment_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
    elemental_functions: &HashSet<String>,
) {
    let StmtKind::Assign { targets, value, .. } = &statement.kind else {
        return;
    };
    let [target] = targets.as_slice() else {
        return;
    };

    let Some(extent) =
        fortran_array_assignment_extent(target, value, array_sizes, array_field_sizes)
    else {
        return;
    };

    let target_is_member = matches!(target.kind, ExprKind::Member { .. });
    let should_lower = contains_fortran_slice(target)
        || expr_is_fortran_elemental_array_call(
            value,
            arrays,
            array_sizes,
            array_field_sizes,
            array_fields,
            array_functions,
            elemental_functions,
        )
        || (target_is_member
            && expr_is_known_fortran_array(
                target,
                array_sizes,
                array_field_sizes,
                array_functions,
            )
            && expr_is_known_fortran_array(value, array_sizes, array_field_sizes, array_functions));
    if !should_lower {
        return;
    }

    let loop_var = "__fortran_array_index";
    let loop_expr = Expression::ident(loop_var);
    let lowered_target = lower_fortran_array_assignment_target(target, &loop_expr);
    let lowered_value = lower_fortran_array_assignment_value(
        value,
        &loop_expr,
        arrays,
        array_sizes,
        array_field_sizes,
        array_fields,
        array_functions,
        elemental_functions,
    );

    *statement = Statement::new(StmtKind::Block(vec![Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(loop_var)],
            value: Expression::int(0),
            by_ref: false,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::ident(loop_var)),
            op: BinOp::Lt,
            right: Box::new(extent),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(loop_var)),
            value: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::ident(loop_var)),
                op: BinOp::Add,
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: vec![Statement::new(StmtKind::Assign {
            targets: vec![lowered_target],
            value: lowered_value,
            by_ref: false,
        })],
    })]));
}

fn fortran_array_assignment_extent(
    target: &Expression,
    value: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    if let Some(extent) = fortran_slice_extent(target) {
        return Some(extent);
    }
    resolve_fortran_array_target_size(target, array_sizes, array_field_sizes)
        .or_else(|| resolve_fortran_array_expr_size(value, array_sizes, array_field_sizes))
}

fn fortran_slice_extent(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Index { object, index, .. } => match &index.kind {
            ExprKind::Slice { lower, upper, .. } => {
                let lower = lower
                    .as_deref()
                    .cloned()
                    .unwrap_or_else(|| Expression::int(0));
                upper.as_deref().cloned().map(|upper| {
                    Expression::new(ExprKind::Binary {
                        left: Box::new(upper),
                        op: BinOp::Sub,
                        right: Box::new(lower),
                    })
                })
            }
            _ => fortran_slice_extent(object),
        },
        _ => None,
    }
}

fn contains_fortran_slice(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Index { object, index, .. } => {
            matches!(index.kind, ExprKind::Slice { .. }) || contains_fortran_slice(object)
        }
        ExprKind::Member { object, .. } => contains_fortran_slice(object),
        _ => false,
    }
}

fn lower_fortran_array_assignment_target(
    target: &Expression,
    loop_index: &Expression,
) -> Expression {
    match &target.kind {
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => match &index.kind {
            ExprKind::Slice { lower, .. } => {
                let base_index = lower
                    .as_deref()
                    .cloned()
                    .unwrap_or_else(|| Expression::int(0));
                Expression::new(ExprKind::Index {
                    object: object.clone(),
                    index: Box::new(Expression::new(ExprKind::Binary {
                        left: Box::new(base_index),
                        op: BinOp::Add,
                        right: Box::new(loop_index.clone()),
                    })),
                    null_safe: *null_safe,
                })
            }
            _ => Expression::new(ExprKind::Index {
                object: Box::new(lower_fortran_array_assignment_target(object, loop_index)),
                index: index.clone(),
                null_safe: *null_safe,
            }),
        },
        _ => Expression::new(ExprKind::Index {
            object: Box::new(target.clone()),
            index: Box::new(loop_index.clone()),
            null_safe: false,
        }),
    }
}

fn lower_fortran_array_assignment_value(
    expr: &Expression,
    loop_index: &Expression,
    arrays: &HashSet<String>,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
    elemental_functions: &HashSet<String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Call {
            callee,
            args,
            optional,
        } if !*optional
            && expr_is_fortran_elemental_array_call(
                expr,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            ) =>
        {
            Expression::new(ExprKind::Call {
                callee: callee.clone(),
                args: args
                    .iter()
                    .map(|arg| Argument {
                        value: lower_fortran_array_assignment_value(
                            &arg.value,
                            loop_index,
                            arrays,
                            array_sizes,
                            array_field_sizes,
                            array_fields,
                            array_functions,
                            elemental_functions,
                        ),
                        name: arg.name.clone(),
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    })
                    .collect(),
                optional: false,
            })
        }
        ExprKind::Ident(_)
        | ExprKind::Member { .. }
        | ExprKind::Slice { .. }
        | ExprKind::Array(_)
        | ExprKind::ArrayTransform { .. }
        | ExprKind::Call { .. } => {
            if matches!(expr.kind, ExprKind::Array(_) | ExprKind::ArrayTransform { .. })
                || matches!(&expr.kind, ExprKind::Member { field, .. } if array_fields.contains(&field.to_ascii_lowercase()))
                || matches!(&expr.kind, ExprKind::Ident(name) if arrays.contains(&name.to_ascii_lowercase()))
                || expr_is_known_fortran_array(
                    expr,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                )
            {
                lower_fortran_array_assignment_target(expr, loop_index)
            } else {
                expr.clone()
            }
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(lower_fortran_array_assignment_value(
                object,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
            index: Box::new(lower_fortran_array_assignment_value(
                index,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
            null_safe: *null_safe,
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(lower_fortran_array_assignment_value(
                left,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
            right: Box::new(lower_fortran_array_assignment_value(
                right,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
        }),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(lower_fortran_array_assignment_value(
                inner,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(lower_fortran_array_assignment_value(
                cond,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
            then: Box::new(lower_fortran_array_assignment_value(
                then,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
            else_: Box::new(lower_fortran_array_assignment_value(
                else_,
                loop_index,
                arrays,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
                elemental_functions,
            )),
        }),
        _ => expr.clone(),
    }
}

fn expr_is_fortran_elemental_array_call(
    expr: &Expression,
    arrays: &HashSet<String>,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
    elemental_functions: &HashSet<String>,
) -> bool {
    let ExprKind::Call {
        callee,
        args,
        optional,
    } = &expr.kind
    else {
        return false;
    };
    if *optional {
        return false;
    }
    let ExprKind::Ident(name) = &callee.kind else {
        return false;
    };
    if !elemental_functions.contains(&name.to_ascii_lowercase()) {
        return false;
    }
    args.iter().any(|arg| {
        matches!(arg.value.kind, ExprKind::Array(_))
            || matches!(&arg.value.kind, ExprKind::Ident(name) if arrays.contains(&name.to_ascii_lowercase()))
            || matches!(&arg.value.kind, ExprKind::Member { field, .. } if array_fields.contains(&field.to_ascii_lowercase()))
            || expr_is_known_fortran_array(&arg.value, array_sizes, array_field_sizes, array_functions)
    })
}

fn lower_fortran_array_return_calls(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_return_calls_with_env(
        body,
        &mut array_sizes,
        &array_field_sizes,
        &array_functions,
        &HashSet::new(),
    );
}

fn lower_fortran_array_return_calls_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
    procedure_params: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_return_call_statement(
            statement,
            array_sizes,
            array_field_sizes,
            array_functions,
            procedure_params,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::Expr(expr) => {
                let mut ignored_ranks = HashMap::new();
                record_fortran_allocate_array_metadata(expr, array_sizes, &mut ignored_ranks, None);
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_return_calls_in_members(
                    members,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                );
            }
            StmtKind::FunctionDecl {
                params,
                body: function_body,
                return_type,
                is_sub,
                ..
            } => {
                let mut nested_arrays = array_sizes.clone();
                let nested_procedure_params =
                    collect_fortran_array_procedure_params(params, array_functions);
                lower_fortran_array_return_calls_with_env(
                    function_body,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &nested_procedure_params,
                );

                if !*is_sub
                    && return_type
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    let result_type = return_type.take();
                    params.push(Param {
                        name: FORTRAN_ARRAY_RESULT_PARAM.to_string(),
                        type_hint: result_type.map(Into::into),
                        default: None,
                        pass_by: PassBy::Out,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                    *is_sub = true;
                    rewrite_fortran_array_return_statements(function_body);
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_return_calls_with_env(
                    std::slice::from_mut(body.as_mut()),
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    then_body,
                    &mut then_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        elif_body,
                        &mut elif_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_return_calls_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        default,
                        &mut default_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    try_body,
                    &mut try_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        finally,
                        &mut finally_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_return_calls_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    params,
                    body,
                    return_type,
                    is_sub,
                    ..
                } = &mut stmt.kind
                else {
                    continue;
                };
                let mut nested_arrays = array_sizes.clone();
                let procedure_params =
                    collect_fortran_array_procedure_params(params, array_functions);
                lower_fortran_array_return_calls_with_env(
                    body,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &procedure_params,
                );
                if !*is_sub
                    && return_type
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    let result_type = return_type.take();
                    params.push(Param {
                        name: FORTRAN_ARRAY_RESULT_PARAM.to_string(),
                        type_hint: result_type.map(Into::into),
                        default: None,
                        pass_by: PassBy::Out,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                    *is_sub = true;
                    rewrite_fortran_array_return_statements(body);
                }
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &HashSet::new(),
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_return_call_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
    procedure_params: &HashSet<String>,
) {
    let StmtKind::Assign { targets, value, .. } = &statement.kind else {
        return;
    };
    let [target] = targets.as_slice() else {
        return;
    };
    if resolve_fortran_array_target_size(target, array_sizes, array_field_sizes).is_none() {
        return;
    }
    let ExprKind::Call {
        callee,
        args,
        optional,
    } = &value.kind
    else {
        return;
    };
    if *optional {
        return;
    }
    let ExprKind::Ident(name) = &callee.kind else {
        return;
    };
    let lower = name.to_ascii_lowercase();
    if !array_functions.contains(&lower) && !procedure_params.contains(&lower) {
        return;
    }

    let mut lowered_args = args.clone();
    lowered_args.push(Argument {
        value: target.clone(),
        name: None,
        by_ref: true,
        spread: false,
    });
    *statement = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: callee.clone(),
        args: lowered_args,
        optional: false,
    })));
}

fn rewrite_fortran_array_return_statements(body: &mut [Statement]) {
    for statement in body.iter_mut() {
        match &mut statement.kind {
            StmtKind::Return(Some(expr)) => {
                *statement = Statement::new(StmtKind::Block(vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(FORTRAN_ARRAY_RESULT_PARAM)],
                        value: expr.clone(),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(None)),
                ]));
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                rewrite_fortran_array_return_statements(stmts)
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                rewrite_fortran_array_return_statements(std::slice::from_mut(body.as_mut()))
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_fortran_array_return_statements(then_body);
                for (_, elif_body) in elifs {
                    rewrite_fortran_array_return_statements(elif_body);
                }
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                rewrite_fortran_array_return_statements(stmts);
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                if let Some(init) = init {
                    rewrite_fortran_array_return_statements(std::slice::from_mut(init.as_mut()));
                }
                rewrite_fortran_array_return_statements(stmts);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    rewrite_fortran_array_return_statements(&mut case.body);
                }
                if let Some(default) = default {
                    rewrite_fortran_array_return_statements(default);
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_fortran_array_return_statements(try_body);
                for catch in catches {
                    rewrite_fortran_array_return_statements(&mut catch.body);
                }
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
                if let Some(finally) = finally {
                    rewrite_fortran_array_return_statements(finally);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_scalar_array_assignments_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_ranks: &mut HashMap<String, usize>,
    array_fields: &HashSet<String>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &mut HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_scalar_array_assignment(
            statement,
            array_sizes,
            array_ranks,
            array_fields,
            array_field_sizes,
            array_field_ranks,
            array_functions,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                    if let Some(rank) = declaration
                        .array_bounds
                        .as_ref()
                        .map(Vec::len)
                        .filter(|rank| *rank > 0)
                    {
                        array_ranks.insert(name.to_ascii_lowercase(), rank);
                    }
                }
            }
            StmtKind::Expr(expr) => {
                record_fortran_allocate_array_metadata(
                    expr,
                    array_sizes,
                    array_ranks,
                    Some(array_field_ranks),
                );
            }
            StmtKind::FunctionDecl {
                params,
                body: function_body,
                ..
            } => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                for param in params.iter() {
                    let array_rank = param
                        .type_hint
                        .as_deref()
                        .map(fortran_type_hint_array_rank)
                        .unwrap_or(0);
                    if array_rank > 0 {
                        nested_ranks.insert(param.name.to_ascii_lowercase(), array_rank);
                    }
                }
                lower_fortran_scalar_array_assignments_with_env(
                    function_body,
                    &mut nested,
                    &mut nested_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut nested,
                    &mut nested_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_scalar_array_assignments_with_env(
                    std::slice::from_mut(body.as_mut()),
                    array_sizes,
                    array_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = array_sizes.clone();
                let mut then_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    then_body,
                    &mut then_arrays,
                    &mut then_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    let mut elif_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        elif_body,
                        &mut elif_arrays,
                        &mut elif_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                if let Some(init) = init {
                    lower_fortran_scalar_array_assignments_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        &mut loop_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    let mut case_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        &mut case_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    let mut default_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        default,
                        &mut default_arrays,
                        &mut default_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = array_sizes.clone();
                let mut try_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    try_body,
                    &mut try_arrays,
                    &mut try_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    let mut catch_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        &mut catch_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    let mut finally_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        finally,
                        &mut finally_arrays,
                        &mut finally_ranks,
                        array_fields,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_scalar_array_assignments_in_members(
                    members,
                    array_sizes,
                    array_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn lower_fortran_scalar_array_assignments_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_fields: &HashSet<String>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &mut HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested,
                    &mut nested_ranks,
                    array_fields,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_scalar_array_assignment(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &mut HashMap<String, usize>,
    _array_fields: &HashSet<String>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Assign { targets, value, .. } => {
            if let Some(rank) = resolve_fortran_array_expr_rank(value, array_ranks, array_field_ranks) {
                for target in targets.iter() {
                    if let Some(key) = fortran_array_target_key(target) {
                        array_ranks.insert(key, rank);
                    }
                }
            }
            if expr_is_known_fortran_array(value, array_sizes, array_field_sizes, array_functions)
                || resolve_fortran_array_expr_rank(value, array_ranks, array_field_ranks).is_some()
            {
                return;
            }
            if !should_broadcast_fortran_array_value(value) {
                return;
            }
            if let Some(rank) = targets.iter().find_map(|target| {
                resolve_fortran_array_target_rank(target, array_ranks, array_field_ranks)
            }) {
                if rank > 1 {
                    if let Some(target) = targets.first() {
                        *value = build_fortran_nested_array_broadcast(
                            target.clone(),
                            rank,
                            value.clone(),
                            0,
                        );
                        return;
                    }
                }
            }
            let Some(size) = targets.iter().find_map(|target| {
                resolve_fortran_array_target_size(target, array_sizes, array_field_sizes)
            }) else {
                return;
            };
            *value = build_fortran_array_fill(size, value.clone());
        }
        _ => {}
    }
}

fn should_broadcast_fortran_array_value(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::Member { .. } => true,
        ExprKind::Unary { expr, .. } => should_broadcast_fortran_array_value(expr),
        ExprKind::Binary { left, right, .. } => {
            should_broadcast_fortran_array_value(left)
                && should_broadcast_fortran_array_value(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            should_broadcast_fortran_array_value(cond)
                && should_broadcast_fortran_array_value(then)
                && should_broadcast_fortran_array_value(else_)
        }
        _ => false,
    }
}

fn resolve_fortran_array_target_size(
    target: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    if let Some(key) = fortran_array_target_key(target) {
        if let Some(size) = array_sizes.get(&key) {
            return Some(size.clone());
        }
    }
    match &target.kind {
        ExprKind::Ident(name) => array_sizes.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Member { field, .. } => {
            array_field_sizes.get(&field.to_ascii_lowercase()).cloned()
        }
        _ => None,
    }
}

fn resolve_fortran_array_target_rank(
    target: &Expression,
    array_ranks: &HashMap<String, usize>,
    array_field_ranks: &HashMap<String, usize>,
) -> Option<usize> {
    if let Some(key) = fortran_array_target_key(target) {
        if let Some(rank) = array_ranks.get(&key) {
            return Some(*rank);
        }
    }
    match &target.kind {
        ExprKind::Ident(name) => array_ranks.get(&name.to_ascii_lowercase()).copied(),
        ExprKind::Member { field, .. } => {
            array_field_ranks.get(&field.to_ascii_lowercase()).copied()
        }
        _ => None,
    }
}

fn expr_is_known_fortran_array(
    expr: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
) -> bool {
    if let Some(key) = fortran_array_target_key(expr) {
        if array_sizes.contains_key(&key) {
            return true;
        }
    }
    match &expr.kind {
        ExprKind::Ident(name) => array_sizes.contains_key(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => {
            array_field_sizes.contains_key(&field.to_ascii_lowercase())
        }
        ExprKind::Array(_) => true,
        ExprKind::Binary { left, right, .. } => {
            expr_is_known_fortran_array(left, array_sizes, array_field_sizes, array_functions)
                || expr_is_known_fortran_array(
                    right,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                )
        }
        ExprKind::Unary { expr: inner, .. } => {
            expr_is_known_fortran_array(inner, array_sizes, array_field_sizes, array_functions)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_is_known_fortran_array(cond, array_sizes, array_field_sizes, array_functions)
                || expr_is_known_fortran_array(
                    then,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                )
                || expr_is_known_fortran_array(
                    else_,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                )
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                name.eq_ignore_ascii_case("Array")
                    || array_functions.contains(&name.to_ascii_lowercase())
            }
            ExprKind::Member { field, .. } => matches!(
                field.to_ascii_lowercase().as_str(),
                "map" | "filter" | "flatmap"
            ),
            _ => false,
        },
        ExprKind::Slice { .. } => true,
        _ => false,
    }
}

fn record_fortran_allocate_array_metadata(
    expr: &Expression,
    array_sizes: &mut HashMap<String, Expression>,
    array_ranks: &mut HashMap<String, usize>,
    array_field_ranks: Option<&mut HashMap<String, usize>>,
) {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return;
    };
    if !name.eq_ignore_ascii_case("allocate") {
        return;
    }

    let mut array_field_ranks = array_field_ranks;
    for arg in args {
        let Some((target, size, rank)) = fortran_allocate_target_metadata(&arg.value) else {
            continue;
        };
        array_sizes.insert(target.clone(), size);
        array_ranks.insert(target, rank);
        if let Some(field_ranks) = array_field_ranks.as_deref_mut() {
            if let Some(field_name) = fortran_member_field_key(&arg.value) {
                field_ranks.insert(field_name, rank);
            }
        }
    }
}

fn fortran_member_field_key(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { field, .. } = &callee.kind else {
        return None;
    };
    Some(field.to_ascii_lowercase())
}

fn fortran_allocate_target_metadata(expr: &Expression) -> Option<(String, Expression, usize)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let target = fortran_array_target_key(callee)?;
    if args.is_empty() {
        return None;
    }
    let rank = args.len();
    let size = args
        .iter()
        .map(|arg| arg.value.clone())
        .reduce(|left, right| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(left),
                right: Box::new(right),
            })
        })?;
    Some((target, size, rank))
}

fn fortran_array_target_key(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.to_ascii_lowercase()),
        ExprKind::Member { object, field, .. } => Some(format!(
            "{}%{}",
            fortran_array_target_key(object)?,
            field.to_ascii_lowercase()
        )),
        _ => None,
    }
}

fn lower_fortran_array_expressions(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_ranks = HashMap::new();
    let mut arrays = HashSet::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_field_ranks = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_field_ranks(body, &mut array_field_ranks);
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_expressions_with_env(
        body,
        &mut array_sizes,
        &mut array_ranks,
        &mut arrays,
        &array_field_sizes,
        &mut array_field_ranks,
        &array_fields,
        &array_functions,
    );
}

fn lower_fortran_array_expressions_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_ranks: &mut HashMap<String, usize>,
    arrays: &mut HashSet<String>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &mut HashMap<String, usize>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_expressions_in_statement(
            statement,
            array_sizes,
            array_ranks,
            array_field_sizes,
            array_field_ranks,
            arrays,
            array_fields,
            array_functions,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    // ⛔ A scalar CHARACTER does NOT belong here. This set drives
                    // the ELEMENTWISE lowerings below — binary, unary and
                    // intrinsic broadcasting — and a `character(len=4) :: c` is
                    // one scalar value, not four. Listing it turned `c == 'zz'`
                    // into a map over the characters, whose non-empty array
                    // result is TRUTHY: every character comparison written
                    // against a bare variable answered `.true.`. A character
                    // ARRAY still arrives through `array_bounds`.
                    if declaration.array_bounds.is_some()
                        || declaration
                            .init
                            .as_ref()
                            .is_some_and(is_array_initializer_expr)
                    {
                        arrays.insert(name.to_ascii_lowercase());
                    }
                    if let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    {
                        array_sizes.insert(name.to_ascii_lowercase(), size);
                    }
                    if let Some(rank) = declaration
                        .array_bounds
                        .as_ref()
                        .map(Vec::len)
                        .filter(|rank| *rank > 0)
                    {
                        array_ranks.insert(name.to_ascii_lowercase(), rank);
                    }
                }
            }
            StmtKind::Expr(expr) => {
                record_fortran_allocate_array_metadata(
                    expr,
                    array_sizes,
                    array_ranks,
                    Some(array_field_ranks),
                );
            }
            StmtKind::FunctionDecl {
                params,
                body,
                name,
                return_type,
                ..
            } => {
                let mut nested = arrays.clone();
                let mut nested_sizes = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                for param in params {
                    let array_rank = param
                        .type_hint
                        .as_deref()
                        .map(fortran_type_hint_array_rank)
                        .unwrap_or(0);
                    if array_rank > 0
                        || param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        nested.insert(param.name.to_ascii_lowercase());
                    }
                    if array_rank > 0 {
                        nested_ranks.insert(param.name.to_ascii_lowercase(), array_rank);
                    }
                }
                let mut nested_functions = array_functions.clone();
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    nested_functions.insert(name.to_ascii_lowercase());
                }
                lower_fortran_array_expressions_with_env(
                    body,
                    &mut nested_sizes,
                    &mut nested_ranks,
                    &mut nested,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    &nested_functions,
                );
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested = arrays.clone();
                let mut nested_sizes = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_array_expressions_with_env(
                    stmts,
                    &mut nested_sizes,
                    &mut nested_ranks,
                    &mut nested,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_array_expressions_with_env(
                    std::slice::from_mut(body.as_mut()),
                    array_sizes,
                    array_ranks,
                    arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_expressions_in_members(
                    members,
                    array_sizes,
                    array_ranks,
                    arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_arrays = arrays.clone();
                let mut then_sizes = array_sizes.clone();
                let mut then_ranks = array_ranks.clone();
                lower_fortran_array_expressions_with_env(
                    then_body,
                    &mut then_sizes,
                    &mut then_ranks,
                    &mut then_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    let mut elif_sizes = array_sizes.clone();
                    let mut elif_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        elif_body,
                        &mut elif_sizes,
                        &mut elif_ranks,
                        &mut elif_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_sizes = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        else_body,
                        &mut else_sizes,
                        &mut else_ranks,
                        &mut else_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_sizes = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                lower_fortran_array_expressions_with_env(
                    stmts,
                    &mut loop_sizes,
                    &mut loop_ranks,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_sizes = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        else_body,
                        &mut else_sizes,
                        &mut else_ranks,
                        &mut else_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_sizes = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                if let Some(init) = init {
                    lower_fortran_array_expressions_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_sizes,
                        &mut loop_ranks,
                        &mut loop_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
                lower_fortran_array_expressions_with_env(
                    stmts,
                    &mut loop_sizes,
                    &mut loop_ranks,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    let mut case_sizes = array_sizes.clone();
                    let mut case_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        &mut case.body,
                        &mut case_sizes,
                        &mut case_ranks,
                        &mut case_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    let mut default_sizes = array_sizes.clone();
                    let mut default_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        default,
                        &mut default_sizes,
                        &mut default_ranks,
                        &mut default_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_arrays = arrays.clone();
                let mut try_sizes = array_sizes.clone();
                let mut try_ranks = array_ranks.clone();
                lower_fortran_array_expressions_with_env(
                    try_body,
                    &mut try_sizes,
                    &mut try_ranks,
                    &mut try_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    let mut catch_sizes = array_sizes.clone();
                    let mut catch_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        &mut catch.body,
                        &mut catch_sizes,
                        &mut catch_ranks,
                        &mut catch_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_sizes = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        else_body,
                        &mut else_sizes,
                        &mut else_ranks,
                        &mut else_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    let mut finally_sizes = array_sizes.clone();
                    let mut finally_ranks = array_ranks.clone();
                    lower_fortran_array_expressions_with_env(
                        finally,
                        &mut finally_sizes,
                        &mut finally_ranks,
                        &mut finally_arrays,
                        array_field_sizes,
                        array_field_ranks,
                        array_fields,
                        array_functions,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_expressions_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &mut HashMap<String, usize>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    params,
                    body,
                    name,
                    return_type,
                    ..
                } = &mut stmt.kind
                else {
                    continue;
                };
                let mut method_arrays = arrays.clone();
                let mut method_sizes = array_sizes.clone();
                let mut method_ranks = array_ranks.clone();
                for param in params.iter() {
                    let array_rank = param
                        .type_hint
                        .as_deref()
                        .map(fortran_type_hint_array_rank)
                        .unwrap_or(0);
                    if array_rank > 0
                        || param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_fortran_string_type_hint)
                    {
                        method_arrays.insert(param.name.to_ascii_lowercase());
                    }
                    if array_rank > 0 {
                        method_ranks.insert(param.name.to_ascii_lowercase(), array_rank);
                    }
                }
                let mut method_functions = array_functions.clone();
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    method_functions.insert(name.to_ascii_lowercase());
                }
                lower_fortran_array_expressions_with_env(
                    body,
                    &mut method_sizes,
                    &mut method_ranks,
                    &mut method_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    &method_functions,
                );
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = arrays.clone();
                let mut nested_sizes = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_array_expressions_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested_sizes,
                    &mut nested_ranks,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_field_ranks,
                    array_fields,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_expressions_in_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            if !is_fortran_allocator_intrinsic_expr(expr) {
                rewrite_fortran_array_expressions_in_expr(
                    expr,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_fortran_array_expressions_in_expr(
                    target,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            rewrite_fortran_array_expressions_in_expr(
                value,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                target,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                value,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        StmtKind::Return(Some(expr)) => rewrite_fortran_array_expressions_in_expr(
            expr,
            array_sizes,
            array_ranks,
            array_field_sizes,
            array_field_ranks,
            arrays,
            array_fields,
            array_functions,
        ),
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_fortran_array_expressions_in_expr(
                    expr,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            if let Some(cause) = cause {
                rewrite_fortran_array_expressions_in_expr(
                    cause,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        StmtKind::If { cond, .. }
        | StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                cond,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_fortran_array_expressions_in_expr(
                    cond,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            if let Some(update) = update {
                rewrite_fortran_array_expressions_in_expr(
                    update,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        StmtKind::ForIn { iter, .. }
        | StmtKind::Using { resource: iter, .. }
        | StmtKind::Lock { expr: iter, .. }
        | StmtKind::Switch { expr: iter, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                iter,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        _ => {}
    }
}

fn rewrite_fortran_array_expressions_in_expr(
    expr: &mut Expression,
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            rewrite_fortran_array_expressions_in_expr(
                left,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                right,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            if let Some(lowered) = lower_fortran_array_binary_expr(
                *op,
                left.as_ref(),
                right.as_ref(),
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            ) {
                *expr = lowered;
            }
        }
        ExprKind::Unary { op, expr: inner } => {
            rewrite_fortran_array_expressions_in_expr(
                inner,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            if let Some(lowered) = lower_fortran_array_unary_expr(
                *op,
                inner.as_ref(),
                array_ranks,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            ) {
                *expr = lowered;
            }
        }
        ExprKind::Await(inner) | ExprKind::YieldFrom(inner) | ExprKind::TypeOf(inner) => {
            rewrite_fortran_array_expressions_in_expr(
                inner,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            )
        }
        // An `ArrayMap` built by an earlier fold — `merge(1, 0, v < 3)` becomes
        // a map whose ARRAY is the mask. Without this arm the traversal stopped
        // at the node and the mask kept its raw `v < 3`, which is a comparison
        // against a whole array and fails in `wasm:js-number.toF64`. The fold
        // was firing correctly; the traversal was not reaching what it built.
        ExprKind::ArrayMap { array, body, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                array,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                body,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_fortran_array_expressions_in_expr(
                cond,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                then,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                else_,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        ExprKind::Member { object, .. } => rewrite_fortran_array_expressions_in_expr(
            object,
            array_sizes,
            array_ranks,
            array_field_sizes,
            array_field_ranks,
            arrays,
            array_fields,
            array_functions,
        ),
        ExprKind::Index { object, index, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                object,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                index,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_fortran_array_expressions_in_expr(
                    lower,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_fortran_array_expressions_in_expr(
                    upper,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            if let Some(step) = step.as_mut() {
                rewrite_fortran_array_expressions_in_expr(
                    step,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_fortran_array_expressions_in_expr(
                callee,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            for arg in args.iter_mut() {
                rewrite_fortran_array_expressions_in_expr(
                    &mut arg.value,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
            if let Some(lowered) = lower_fortran_array_intrinsic_expr(
                callee.as_ref(),
                args,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            ) {
                *expr = lowered;
            }
        }
        ExprKind::New { class, args } => {
            rewrite_fortran_array_expressions_in_expr(
                class,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            for arg in args.iter_mut() {
                rewrite_fortran_array_expressions_in_expr(
                    &mut arg.value,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_fortran_array_expressions_in_expr(
                target,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
            rewrite_fortran_array_expressions_in_expr(
                value,
                array_sizes,
                array_ranks,
                array_field_sizes,
                array_field_ranks,
                arrays,
                array_fields,
                array_functions,
            );
        }
        // PACK/UNPACK/RESHAPE/MERGE carry ORDINARY expressions as arguments —
        // `merge(a, b, a < b)`'s mask is an elementwise comparison that still
        // has to be repaired. Without this arm the args were skipped entirely
        // and the mask reached the runtime as a scalar comparison of two
        // arrays. Same shape as every other missing-traversal-arm bug: a node
        // kind needs an arm in EVERY pass that walks expressions.
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args {
                rewrite_fortran_array_expressions_in_expr(
                    arg,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = item.key.as_mut() {
                    rewrite_fortran_array_expressions_in_expr(
                        key,
                        array_sizes,
                        array_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        arrays,
                        array_fields,
                        array_functions,
                    );
                }
                rewrite_fortran_array_expressions_in_expr(
                    &mut item.value,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_fortran_array_expressions_in_expr(
                    item,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    arrays,
                    array_fields,
                    array_functions,
                );
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_fortran_array_expressions_in_expr(
                            key,
                            array_sizes,
                            array_ranks,
                            array_field_sizes,
                            array_field_ranks,
                            arrays,
                            array_fields,
                            array_functions,
                        );
                        rewrite_fortran_array_expressions_in_expr(
                            value,
                            array_sizes,
                            array_ranks,
                            array_field_sizes,
                            array_field_ranks,
                            arrays,
                            array_fields,
                            array_functions,
                        );
                    }
                    ObjectProperty::Spread(expr) => rewrite_fortran_array_expressions_in_expr(
                        expr,
                        array_sizes,
                        array_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        arrays,
                        array_fields,
                        array_functions,
                    ),
                    ObjectProperty::Shorthand(_)
                    | ObjectProperty::Method { .. }
                    | ObjectProperty::Accessor { .. } => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    rewrite_fortran_array_expressions_in_expr(
                        expr,
                        array_sizes,
                        array_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        arrays,
                        array_fields,
                        array_functions,
                    );
                }
            }
        }
        _ => {}
    }
}

fn lower_fortran_array_binary_expr(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    _array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    _array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Option<Expression> {
    // Comparisons broadcast exactly as arithmetic does — `values >= 2` is a
    // LOGICAL array, which is what `where (values >= 2)` and `count(values <= 4)`
    // are built out of. Leaving them off this list did not make them scalar; it
    // made them a comparison against a whole array, which reached
    // `wasm:js-number.toF64` and failed there.
    if !matches!(
        op,
        BinOp::Add
            | BinOp::Sub
            | BinOp::Mul
            | BinOp::Div
            | BinOp::Pow
            | BinOp::Lt
            | BinOp::Gt
            | BinOp::LtEq
            | BinOp::GtEq
            | BinOp::Eq
            | BinOp::NotEq
            // `.and.`/`.or.` combine two LOGICAL arrays elementwise, which is
            // what `ELSEWHERE` narrows its mask with.
            | BinOp::And
            | BinOp::Or
            | BinOp::Eqv
    ) {
        return None;
    }
    let left_is_array = is_known_fortran_array_expr(left, arrays, array_fields, array_functions);
    let right_is_array = is_known_fortran_array_expr(right, arrays, array_fields, array_functions);
    if !left_is_array && !right_is_array {
        return None;
    }

    let rank = [left, right]
        .into_iter()
        .filter(|expr| is_known_fortran_array_expr(expr, arrays, array_fields, array_functions))
        .filter_map(|expr| resolve_fortran_array_expr_rank(expr, array_ranks, array_field_ranks))
        .max()
        .unwrap_or(1);

    if rank > 1 {
        return Some(build_fortran_nested_array_binary_expr(
            op,
            left.clone(),
            right.clone(),
            rank,
            0,
            left_is_array,
            right_is_array,
        ));
    }

    let item_name = "__fortran_array_item";
    let index_name = "__fortran_array_index";
    let lambda_body = if left_is_array && right_is_array {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(Expression::ident(item_name)),
            right: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(right.clone()),
                index: Box::new(Expression::ident(index_name)),
                null_safe: false,
            })),
        })
    } else if left_is_array {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(Expression::ident(item_name)),
            right: Box::new(right.clone()),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left.clone()),
            right: Box::new(Expression::ident(item_name)),
        })
    };

    Some(build_fortran_array_map(
        if left_is_array {
            left.clone()
        } else {
            right.clone()
        },
        lambda_body,
        left_is_array && right_is_array,
        item_name,
        index_name,
    ))
}

fn lower_fortran_array_unary_expr(
    op: UnaryOp,
    value: &Expression,
    array_ranks: &HashMap<String, usize>,
    array_field_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Option<Expression> {
    // `.not.` on a LOGICAL array negates each element — the form `ELSEWHERE`
    // needs to exclude what an earlier clause already claimed.
    if !matches!(op, UnaryOp::Neg | UnaryOp::Pos | UnaryOp::Not) {
        return None;
    }
    if !is_known_fortran_array_expr(value, arrays, array_fields, array_functions) {
        return None;
    }

    let rank = resolve_fortran_array_expr_rank(value, array_ranks, array_field_ranks).unwrap_or(1);
    Some(build_fortran_nested_array_unary_expr(
        op,
        value.clone(),
        rank,
        0,
    ))
}

fn lower_fortran_array_intrinsic_expr(
    callee: &Expression,
    args: &[Argument],
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let positional_args = args
        .iter()
        .filter(|arg| arg.name.is_none())
        .collect::<Vec<_>>();
    let array_expr = &positional_args.first()?.value;
    if !is_known_fortran_array_expr(array_expr, arrays, array_fields, array_functions) {
        return None;
    }

    let lowered = name.to_ascii_lowercase();
    let rank =
        resolve_fortran_array_expr_rank(array_expr, array_ranks, array_field_ranks).unwrap_or(1);
    match lowered.as_str() {
        "real" | "int" | "dble" if !positional_args.is_empty() => {
            Some(build_fortran_nested_array_intrinsic_call_with_args(
                &lowered,
                array_expr.clone(),
                positional_args
                    .iter()
                    .skip(1)
                    .map(|arg| Argument::positional(arg.value.clone()))
                    .collect(),
                rank,
                0,
            ))
        }
        "maxloc" | "minloc" => {
            let kind = if lowered == "maxloc" { "max" } else { "min" };
            let mask_expr = args.iter().find(|arg| {
                arg.name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case("mask"))
            });
            let dim_arg = args
                .iter()
                .find(|arg| {
                    arg.name
                        .as_deref()
                        .is_some_and(|name| name.eq_ignore_ascii_case("dim"))
                })
                .or_else(|| positional_args.get(1).copied());
            let scalar_loc = if let Some(mask_arg) = mask_expr {
                if rank != 1 {
                    return None;
                }
                build_fortran_masked_rank1_loc_expr(
                    kind,
                    array_expr.clone(),
                    mask_arg.value.clone(),
                )
            } else if let Some(arg) = dim_arg {
                if fortran_dim_is_one(&arg.value) {
                    if rank == 1 {
                        build_fortran_rank1_loc_expr(kind, array_expr.clone())
                    } else {
                        build_fortran_nested_array_loc_expr(kind, array_expr.clone(), rank, 0)
                    }
                } else if fortran_dim_is_two(&arg.value) && rank == 2 {
                    build_fortran_nested_array_loc_expr_dim2(kind, array_expr.clone(), 0)
                } else {
                    return None;
                }
            } else if rank != 1 {
                return None;
            } else {
                build_fortran_rank1_loc_expr(kind, array_expr.clone())
            };
            match dim_arg {
                Some(arg) if fortran_dim_is_one(&arg.value) && rank == 1 => Some(scalar_loc),
                Some(_) if rank > 1 => Some(scalar_loc),
                None if rank == 1 => Some(Expression::new(ExprKind::Array(vec![ArrayElement {
                    key: None,
                    value: scalar_loc,
                    spread: false,
                    by_ref: false,
                }]))),
                _ => None,
            }
        }
        "findloc" => {
            let value_expr = positional_args.get(1)?.value.clone();
            let back = args
                .iter()
                .find(|arg| {
                    arg.name
                        .as_deref()
                        .is_some_and(|name| name.eq_ignore_ascii_case("back"))
                })
                .is_some_and(|arg| fortran_logical_is_true(&arg.value));
            let mask_expr = args.iter().find(|arg| {
                arg.name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case("mask"))
            });
            let dim_arg = args.iter().find(|arg| {
                arg.name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case("dim"))
            });
            if dim_arg.is_some() || rank > 1 {
                return None;
            }
            let loc_expr = build_fortran_findloc_expr(
                array_expr.clone(),
                value_expr,
                back,
                mask_expr.map(|arg| arg.value.clone()),
            );
            Some(Expression::new(ExprKind::Array(vec![ArrayElement {
                key: None,
                value: loc_expr,
                spread: false,
                by_ref: false,
            }])))
        }
        // `maxval(a, mask=m)` and friends — the mask used to be dropped on the
        // floor here, so they answered for the whole array. Rank 1 only: a
        // masked reduction of a rank-2 array needs the mask walked in the same
        // nesting, which this shape does not express.
        "maxval" | "minval" | "sum" | "product"
            if rank == 1
                && args.iter().any(|arg| {
                    arg.name
                        .as_deref()
                        .is_some_and(|name| name.eq_ignore_ascii_case("mask"))
                }) =>
        {
            let mask = args.iter().find_map(|arg| {
                arg.name
                    .as_deref()
                    .filter(|name| name.eq_ignore_ascii_case("mask"))
                    .map(|_| arg.value.clone())
            })?;
            let kind = match lowered.as_str() {
                "maxval" => "max",
                "minval" => "min",
                other => other,
            };
            let masked = build_fortran_masked_neutral_map(kind, array_expr.clone(), mask);
            Some(match kind {
                "product" => build_fortran_product_expr(masked),
                _ => build_fortran_array_reduction(kind, masked, 0),
            })
        }
        _ if args.len() != 1 || args[0].name.is_some() => None,
        "size" => Some(if args.len() == 1 {
            resolve_fortran_array_expr_size(array_expr, array_sizes, array_field_sizes)
                .unwrap_or_else(|| build_fortran_nested_array_size_expr(array_expr.clone(), rank, 0))
        } else if rank > 1 {
            build_fortran_nested_array_size_expr(array_expr.clone(), rank, 0)
        } else {
            resolve_fortran_array_expr_size(array_expr, array_sizes, array_field_sizes)
                .unwrap_or_else(|| {
                    Expression::new(ExprKind::Member {
                        object: Box::new(array_expr.clone()),
                        field: "length".to_string(),
                        null_safe: false,
                    })
                })
        }),
        "sum" => Some(build_fortran_nested_array_reduction(
            "sum",
            array_expr.clone(),
            rank,
            0,
        )),
        // The BIT reductions (F2008 13.7.61-63): `iall` folds with AND, `iany`
        // with OR, `iparity` with XOR. Rank-aware through the same nested
        // reduction `sum` uses, so a rank-2 argument folds rows then elements
        // instead of ANDing arrays together.
        "iall" | "iany" | "iparity" => Some(build_fortran_nested_array_reduction(
            &lowered,
            array_expr.clone(),
            rank,
            0,
        )),
        "minval" => Some(build_fortran_nested_array_reduction(
            "min",
            array_expr.clone(),
            rank,
            0,
        )),
        "maxval" => Some(build_fortran_nested_array_reduction(
            "max",
            array_expr.clone(),
            rank,
            0,
        )),
        "abs" | "acos" | "asin" | "atan" | "cos" | "exp" | "log" | "sin" | "sqrt" | "tan" => Some(
            build_fortran_nested_array_intrinsic_call(&lowered, array_expr.clone(), rank, 0),
        ),
        _ => return None,
    }
}

fn fortran_dim_is_one(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(1)) => true,
        ExprKind::Lit(Literal::Float(value)) => *value == 1.0,
        _ => false,
    }
}

fn fortran_dim_is_two(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(2)) => true,
        ExprKind::Lit(Literal::Float(value)) => *value == 2.0,
        _ => false,
    }
}

fn fortran_logical_is_true(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Bool(true)) => true,
        ExprKind::Ident(name) => name.eq_ignore_ascii_case(".true."),
        _ => false,
    }
}

fn fortran_index_to_loc(index_expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(index_expr.clone()),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(Expression::int(0)),
        else_: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(index_expr),
            right: Box::new(Expression::int(1)),
        })),
    })
}

fn build_fortran_rank1_loc_expr(kind: &str, array_expr: Expression) -> Expression {
    let target_value = build_fortran_array_reduction(kind, array_expr.clone(), 0);
    fortran_index_to_loc(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_expr),
            field: "indexOf".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(target_value)],
        optional: false,
    }))
}

fn fortran_expr_is_true(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr),
            right: Box::new(Expression::int(1)),
        })),
    })
}

fn build_fortran_findloc_expr(
    array_expr: Expression,
    value_expr: Expression,
    back: bool,
    mask_expr: Option<Expression>,
) -> Expression {
    let index_expr = if let Some(mask) = mask_expr {
        let item_name = "__fortran_findloc_item";
        let idx_name = "__fortran_findloc_idx";
        let predicate = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(fortran_expr_is_true(Expression::new(ExprKind::Index {
                object: Box::new(mask),
                index: Box::new(Expression::ident(idx_name)),
                null_safe: false,
            }))),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::ident(item_name)),
                right: Box::new(value_expr),
            })),
        });
        let method = if back { "findLastIndex" } else { "findIndex" };
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(array_expr),
                field: method.to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
                params: vec![
                    Param {
                        name: item_name.to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    },
                    Param {
                        name: idx_name.to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    },
                ],
                body: LambdaBody::Expr(Box::new(predicate)),
                is_async: false,
                captures: Vec::new(),
            }))],
            optional: false,
        })
    } else if back {
        let item_name = "__fortran_findloc_item";
        let predicate = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::ident(item_name)),
            right: Box::new(value_expr),
        });
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(array_expr),
                field: "findLastIndex".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
                params: vec![Param {
                    name: item_name.to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(predicate)),
                is_async: false,
                captures: Vec::new(),
            }))],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(array_expr),
                field: "indexOf".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(value_expr)],
            optional: false,
        })
    };
    fortran_index_to_loc(index_expr)
}

/// A masked array flattened to one the plain reduction can walk: each element
/// the mask excludes becomes the value that cannot change the answer.
///
/// `max`/`min` use a sentinel far outside any plausible operand, `sum` uses 0
/// and `product` uses 1 — the identity of each operation. This is the shape
/// `maxloc(a, mask=m)` already used; naming it lets `maxval`/`minval`/`sum`/
/// `product` reduce the same array instead of ignoring their mask.
fn build_fortran_masked_neutral_map(
    kind: &str,
    array_expr: Expression,
    mask_expr: Expression,
) -> Expression {
    let item_name = "__fortran_masked_loc_item";
    let idx_name = "__fortran_masked_loc_idx";
    let neutral = match kind {
        "max" => Expression::int(-1_000_000_000),
        "min" => Expression::int(1_000_000_000),
        "product" => Expression::int(1),
        _ => Expression::int(0),
    };
    build_fortran_typed_array_map(
        array_expr,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(fortran_expr_is_true(Expression::new(ExprKind::Index {
                object: Box::new(mask_expr),
                index: Box::new(Expression::ident(idx_name)),
                null_safe: false,
            }))),
            then: Box::new(Expression::ident(item_name)),
            else_: Box::new(neutral),
        }),
        true,
        item_name,
        idx_name,
        None,
    )
}

fn build_fortran_masked_rank1_loc_expr(
    kind: &str,
    array_expr: Expression,
    mask_expr: Expression,
) -> Expression {
    build_fortran_rank1_loc_expr(
        kind,
        build_fortran_masked_neutral_map(kind, array_expr, mask_expr),
    )
}

fn build_fortran_nested_array_loc_expr(
    kind: &str,
    array_expr: Expression,
    rank: usize,
    depth: usize,
) -> Expression {
    if rank <= 1 {
        return build_fortran_rank1_loc_expr(kind, array_expr);
    }

    let item_name = format!("__fortran_{kind}_slice_{depth}");
    let index_name = format!("__fortran_{kind}_slice_idx_{depth}");
    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_loc_expr(
            kind,
            Expression::ident(&item_name),
            rank - 1,
            depth + 1,
        ),
        false,
        &item_name,
        &index_name,
    )
}

fn build_fortran_nested_array_loc_expr_dim2(
    kind: &str,
    array_expr: Expression,
    depth: usize,
) -> Expression {
    let item_name = format!("__fortran_{kind}_row_{depth}");
    let index_name = format!("__fortran_{kind}_row_idx_{depth}");
    build_fortran_array_map(
        array_expr,
        build_fortran_rank1_loc_expr(kind, Expression::ident(&item_name)),
        false,
        &item_name,
        &index_name,
    )
}

fn build_fortran_nested_array_binary_expr(
    op: BinOp,
    left: Expression,
    right: Expression,
    rank: usize,
    depth: usize,
    left_is_array: bool,
    right_is_array: bool,
) -> Expression {
    let item_name = format!("__fortran_array_item_{depth}");
    let index_name = format!("__fortran_array_index_{depth}");

    if rank <= 1 {
        let lambda_body = if left_is_array && right_is_array {
            Expression::new(ExprKind::Binary {
                op,
                left: Box::new(Expression::ident(&item_name)),
                right: Box::new(Expression::new(ExprKind::Index {
                    object: Box::new(right.clone()),
                    index: Box::new(Expression::ident(&index_name)),
                    null_safe: false,
                })),
            })
        } else if left_is_array {
            Expression::new(ExprKind::Binary {
                op,
                left: Box::new(Expression::ident(&item_name)),
                right: Box::new(right.clone()),
            })
        } else {
            Expression::new(ExprKind::Binary {
                op,
                left: Box::new(left.clone()),
                right: Box::new(Expression::ident(&item_name)),
            })
        };

        return build_fortran_array_map(
            if left_is_array {
                left.clone()
            } else {
                right.clone()
            },
            lambda_body,
            left_is_array && right_is_array,
            &item_name,
            &index_name,
        );
    }

    let next_left = if left_is_array {
        Expression::ident(&item_name)
    } else {
        left.clone()
    };
    let next_right = if right_is_array {
        if left_is_array {
            Expression::new(ExprKind::Index {
                object: Box::new(right.clone()),
                index: Box::new(Expression::ident(&index_name)),
                null_safe: false,
            })
        } else {
            Expression::ident(&item_name)
        }
    } else {
        right.clone()
    };

    build_fortran_array_map(
        if left_is_array { left } else { right },
        build_fortran_nested_array_binary_expr(
            op,
            next_left,
            next_right,
            rank - 1,
            depth + 1,
            left_is_array,
            right_is_array,
        ),
        left_is_array && right_is_array,
        &item_name,
        &index_name,
    )
}

fn build_fortran_nested_array_reduction(
    kind: &str,
    array_expr: Expression,
    rank: usize,
    depth: usize,
) -> Expression {
    if rank <= 1 {
        return build_fortran_array_reduction(kind, array_expr, depth);
    }

    let item_name = format!("__fortran_{}_item_{depth}", kind);
    let mapped = build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_reduction(
            kind,
            Expression::ident(&item_name),
            rank - 1,
            depth + 1,
        ),
        false,
        &item_name,
        &format!("__fortran_{}_index_{depth}", kind),
    );
    build_fortran_array_reduction(kind, mapped, depth)
}

fn build_fortran_nested_array_unary_expr(
    op: UnaryOp,
    array_expr: Expression,
    rank: usize,
    depth: usize,
) -> Expression {
    let item_name = format!("__fortran_unary_item_{depth}");
    if rank <= 1 {
        return build_fortran_array_map(
            array_expr,
            Expression::new(ExprKind::Unary {
                op,
                expr: Box::new(Expression::ident(&item_name)),
            }),
            false,
            &item_name,
            &format!("__fortran_unary_index_{depth}"),
        );
    }

    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_unary_expr(
            op,
            Expression::ident(&item_name),
            rank - 1,
            depth + 1,
        ),
        false,
        &item_name,
        &format!("__fortran_unary_index_{depth}"),
    )
}

fn build_fortran_nested_array_intrinsic_call(
    name: &str,
    array_expr: Expression,
    rank: usize,
    depth: usize,
) -> Expression {
    let item_name = format!("__fortran_intrinsic_item_{depth}");
    if rank <= 1 {
        return build_fortran_array_map(
            array_expr,
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(name)),
                args: vec![Argument::positional(Expression::ident(&item_name))],
                optional: false,
            }),
            false,
            &item_name,
            &format!("__fortran_intrinsic_index_{depth}"),
        );
    }

    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_intrinsic_call(
            name,
            Expression::ident(&item_name),
            rank - 1,
            depth + 1,
        ),
        false,
        &item_name,
        &format!("__fortran_intrinsic_index_{depth}"),
    )
}

fn build_fortran_nested_array_intrinsic_call_with_args(
    name: &str,
    array_expr: Expression,
    extra_args: Vec<Argument>,
    rank: usize,
    depth: usize,
) -> Expression {
    let item_name = format!("__fortran_intrinsic_item_{depth}");
    let index_name = format!("__fortran_intrinsic_index_{depth}");
    let call_body = {
        let mut call_args = Vec::with_capacity(extra_args.len() + 1);
        call_args.push(Argument::positional(Expression::ident(&item_name)));
        call_args.extend(extra_args.iter().cloned());
        lower_intrinsic_expr_call(&Expression::ident(name), &call_args).unwrap_or_else(|| {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(name)),
                args: call_args,
                optional: false,
            })
        })
    };

    if rank <= 1 {
        return build_fortran_array_map(array_expr, call_body, false, &item_name, &index_name);
    }

    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_intrinsic_call_with_args(
            name,
            Expression::ident(&item_name),
            extra_args,
            rank - 1,
            depth + 1,
        ),
        false,
        &item_name,
        &index_name,
    )
}

/// `REDUCE(ARRAY, OPERATION [, MASK] [, IDENTITY])` — F2018 16.9.161.
///
/// OPERATION is a PURE FUNCTION of two arguments, not an operator: `operator(+)`
/// is a generic-spec and cannot be an actual argument at all. The fold is the
/// ordinary `reduce` the array reductions already use, with the user's function
/// as the reducer body.
///
/// `MASK` selects first, which is exactly `PACK` — the shared ranked-array node
/// rather than a second filtering path. Without a mask a ranked array is
/// FLATTENED, because a `REDUCE` with no `DIM` folds every element whatever the
/// rank, and reducing a rank-2 nest would otherwise hand the function ROWS.
fn build_fortran_reduce_call(
    array: Expression,
    function: &str,
    mask: Option<Expression>,
    identity: Option<Expression>,
) -> Expression {
    let acc_name = "__fortran_reduce_acc";
    let item_name = "__fortran_reduce_item";
    let source = match mask {
        Some(mask) => Expression::new(ExprKind::ArrayTransform {
            op: ArrayTransformOp::PackMask,
            args: vec![array, mask],
            order: ArrayTraversalOrder::ColumnMajor,
        }),
        None => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(array),
                field: "flat".to_string(),
                null_safe: false,
            })),
            args: Vec::new(),
            optional: false,
        }),
    };
    let param = |name: &str| Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    let body = fortran_call(
        function,
        vec![
            Expression::ident(acc_name),
            Expression::ident(item_name),
        ],
    );
    let mut args = vec![Argument::positional(Expression::new(ExprKind::Lambda {
        params: vec![param(acc_name), param(item_name)],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: Vec::new(),
    }))];
    if let Some(identity) = identity {
        args.push(Argument::positional(identity));
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(source),
            field: "reduce".to_string(),
            null_safe: false,
        })),
        args,
        optional: false,
    })
}

fn build_fortran_array_reduction(kind: &str, array_expr: Expression, depth: usize) -> Expression {
    let acc_name = format!("__fortran_{}_acc_{depth}", kind);
    let item_name = format!("__fortran_{}_item_{depth}", kind);
    let reducer_body = match kind {
        "sum" => Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&acc_name)),
            right: Box::new(Expression::ident(&item_name)),
        }),
        "min" | "max" => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(kind)),
            args: vec![
                Argument::positional(Expression::ident(&acc_name)),
                Argument::positional(Expression::ident(&item_name)),
            ],
            optional: false,
        }),
        // `iall`/`iany`/`iparity` fold with AND/OR/XOR. Each needs its IDENTITY
        // as the seed below, or `reduce` starts from element 0 and an empty
        // array throws instead of answering the identity.
        "iall" | "iany" | "iparity" => Expression::new(ExprKind::Binary {
            op: match kind {
                "iall" => BinOp::BitAnd,
                "iany" => BinOp::BitOr,
                _ => BinOp::BitXor,
            },
            left: Box::new(Expression::ident(&acc_name)),
            right: Box::new(Expression::ident(&item_name)),
        }),
        _ => array_expr.clone(),
    };

    let mut args = vec![Argument::positional(Expression::new(ExprKind::Lambda {
        params: vec![
            Param {
                name: acc_name,
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            },
            Param {
                name: item_name,
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            },
        ],
        body: LambdaBody::Expr(Box::new(reducer_body)),
        is_async: false,
        captures: Vec::new(),
    }))];
    // The reduction's IDENTITY. `iall` seeds all-ones so the first AND is a
    // no-op; `iany`/`iparity` seed zero. `min`/`max` deliberately have none —
    // there is no representable identity, so they start from element 0.
    if let Some(identity) = match kind {
        "sum" | "iany" | "iparity" => Some(0),
        "iall" => Some(-1),
        _ => None,
    } {
        args.push(Argument::positional(Expression::int(identity)));
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_expr),
            field: "reduce".to_string(),
            null_safe: false,
        })),
        args,
        optional: false,
    })
}

fn build_fortran_array_map(
    array_expr: Expression,
    body: Expression,
    include_index: bool,
    item_name: &str,
    index_name: &str,
) -> Expression {
    build_fortran_typed_array_map(array_expr, body, include_index, item_name, index_name, None)
}

fn build_fortran_typed_array_map(
    array_expr: Expression,
    body: Expression,
    include_index: bool,
    item_name: &str,
    index_name: &str,
    item_type_hint: Option<String>,
) -> Expression {
    let mut params = vec![Param {
        name: item_name.to_string(),
        type_hint: item_type_hint.map(Into::into),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }];
    if include_index {
        params.push(Param {
            name: index_name.to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }
    Expression::new(ExprKind::ArrayMap {
        array: Box::new(array_expr),
        params,
        body: Box::new(body),
    })
}

fn is_known_fortran_array_expr(
    expr: &Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => arrays.contains(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_fields.contains(&field.to_ascii_lowercase()),
        ExprKind::Array(_) | ExprKind::Slice { .. } => true,
        // The NODE states where its shape comes from; this only has to ask.
        // `PACK` yields a vector whatever went in (`None`); the others wear the
        // shape of one argument, so a `MERGE` under a scalar mask is a SCALAR.
        ExprKind::ArrayTransform { op, args, .. } => match op.shape_source_arg() {
            None => true,
            Some(index) => args.get(index).is_some_and(|arg| {
                is_known_fortran_array_expr(arg, arrays, array_fields, array_functions)
            }),
        },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                matches!(name.to_ascii_lowercase().as_str(), "array")
                    || array_functions.contains(&name.to_ascii_lowercase())
            }
            ExprKind::Member { field, .. } => matches!(
                field.to_ascii_lowercase().as_str(),
                "map" | "filter" | "flatmap"
            ),
            _ => false,
        },
        _ => false,
    }
}

fn resolve_fortran_array_expr_rank(
    expr: &Expression,
    array_ranks: &HashMap<String, usize>,
    array_field_ranks: &HashMap<String, usize>,
) -> Option<usize> {
    if let Some(key) = fortran_array_target_key(expr) {
        if let Some(rank) = array_ranks.get(&key) {
            return Some(*rank);
        }
    }

    match &expr.kind {
        ExprKind::Ident(name) => array_ranks.get(&name.to_ascii_lowercase()).copied(),
        ExprKind::Member { field, .. } => {
            array_field_ranks.get(&field.to_ascii_lowercase()).copied()
        }
        ExprKind::Array(items) => Some(
            items
                .first()
                .and_then(|item| {
                    resolve_fortran_array_expr_rank(&item.value, array_ranks, array_field_ranks)
                })
                .unwrap_or(0)
                + 1,
        ),
        ExprKind::Binary { left, right, .. } => {
            resolve_fortran_array_expr_rank(left, array_ranks, array_field_ranks)
                .or_else(|| resolve_fortran_array_expr_rank(right, array_ranks, array_field_ranks))
        }
        ExprKind::Unary { expr: inner, .. } => {
            resolve_fortran_array_expr_rank(inner, array_ranks, array_field_ranks)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            resolve_fortran_array_expr_rank(cond, array_ranks, array_field_ranks)
                .or_else(|| resolve_fortran_array_expr_rank(then, array_ranks, array_field_ranks))
                .or_else(|| resolve_fortran_array_expr_rank(else_, array_ranks, array_field_ranks))
        }
        ExprKind::Slice { .. } => Some(1),
        ExprKind::Index { object, index, .. } => match &index.kind {
            ExprKind::Slice { .. } => Some(1),
            _ => resolve_fortran_array_expr_rank(object, array_ranks, array_field_ranks)
                .and_then(|rank| rank.checked_sub(1))
                .filter(|rank| *rank > 0),
        },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Member { object, field, .. }
                if matches!(
                    field.to_ascii_lowercase().as_str(),
                    "map" | "filter" | "flatmap"
                ) =>
            {
                resolve_fortran_array_expr_rank(object, array_ranks, array_field_ranks)
            }
            _ => None,
        },
        _ => None,
    }
}

fn build_fortran_array_fill(size: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![Argument::positional(size), Argument::positional(value)],
        optional: false,
    })
}

fn build_fortran_nested_array_size_expr(
    array_expr: Expression,
    rank: usize,
    depth: usize,
) -> Expression {
    if rank <= 1 {
        return Expression::new(ExprKind::Member {
            object: Box::new(array_expr),
            field: "length".to_string(),
            null_safe: false,
        });
    }

    let item_name = format!("__fortran_size_item_{depth}");
    let mapped_sizes = build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_size_expr(Expression::ident(&item_name), rank - 1, depth + 1),
        false,
        &item_name,
        &format!("__fortran_size_index_{depth}"),
    );
    build_fortran_array_reduction("sum", mapped_sizes, depth)
}

fn build_fortran_nested_array_broadcast(
    array_expr: Expression,
    rank: usize,
    value: Expression,
    depth: usize,
) -> Expression {
    if rank <= 1 {
        return build_fortran_array_map(
            array_expr,
            value,
            false,
            &format!("__fortran_broadcast_item_{depth}"),
            &format!("__fortran_broadcast_index_{depth}"),
        );
    }

    let item_name = format!("__fortran_broadcast_item_{depth}");
    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_broadcast(
            Expression::ident(&item_name),
            rank - 1,
            value,
            depth + 1,
        ),
        false,
        &item_name,
        &format!("__fortran_broadcast_index_{depth}"),
    )
}

fn collect_fortran_array_field_ranks(
    body: &[Statement],
    array_field_ranks: &mut HashMap<String, usize>,
) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_ranks_in_members(members, array_field_ranks);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_field_ranks_in_members(members, array_field_ranks);
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_fortran_array_field_ranks(body, array_field_ranks)
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_ranks_in_members(
    members: &[ClassMember],
    array_field_ranks: &mut HashMap<String, usize>,
) {
    for member in members {
        match member {
            ClassMember::Field {
                name, array_bounds, ..
            } => {
                if let Some(rank) = array_bounds.as_ref().map(Vec::len).filter(|rank| *rank > 0) {
                    array_field_ranks.insert(name.to_ascii_lowercase(), rank);
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_ranks(
                    std::slice::from_ref(stmt.as_ref()),
                    array_field_ranks,
                );
            }
            _ => {}
        }
    }
}

fn bounds_total_size_expr(bounds: &[Expression]) -> Option<Expression> {
    let mut iter = bounds.iter().cloned();
    let first = iter.next()?;
    Some(iter.fold(first, |acc, bound| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(acc),
            right: Box::new(bound),
        })
    }))
}

fn array_init_size_expr(init: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &init.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !name.eq_ignore_ascii_case("Array") {
        return None;
    }
    args.first().map(|arg| arg.value.clone())
}

fn rewrite_array_subscripts_in_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    char_vars: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            let preserve_intrinsic_args = matches!(
                &expr.kind,
                ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("allocate") || name.eq_ignore_ascii_case("deallocate"))
            );
            if !preserve_intrinsic_args {
                rewrite_array_subscripts_in_expr(expr, arrays, char_vars, callables, array_fields);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_array_subscripts_in_expr(
                    target,
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
            rewrite_array_subscripts_in_expr(value, arrays, char_vars, callables, array_fields);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_array_subscripts_in_expr(target, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(value, arrays, char_vars, callables, array_fields);
        }
        StmtKind::Return(Some(expr)) => {
            rewrite_array_subscripts_in_expr(expr, arrays, char_vars, callables, array_fields)
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_array_subscripts_in_expr(expr, arrays, char_vars, callables, array_fields);
            }
            if let Some(cause) = cause {
                rewrite_array_subscripts_in_expr(cause, arrays, char_vars, callables, array_fields);
            }
        }
        StmtKind::If { cond, .. }
        | StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. } => {
            rewrite_array_subscripts_in_expr(cond, arrays, char_vars, callables, array_fields)
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_array_subscripts_in_expr(cond, arrays, char_vars, callables, array_fields);
            }
            if let Some(update) = update {
                rewrite_array_subscripts_in_expr(
                    update,
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
        }
        StmtKind::ForIn { iter, .. }
        | StmtKind::Using { resource: iter, .. }
        | StmtKind::Lock { expr: iter, .. }
        | StmtKind::Switch { expr: iter, .. } => {
            rewrite_array_subscripts_in_expr(iter, arrays, char_vars, callables, array_fields);
        }
        _ => {}
    }
}

fn rewrite_array_subscripts_in_expr(
    expr: &mut Expression,
    arrays: &HashSet<String>,
    char_vars: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_array_subscripts_in_expr(left, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(right, arrays, char_vars, callables, array_fields);
        }
        // `real(v(1))` folds to a CAST while the expression is being walked, and
        // a subscript inside one still has to be normalised. Without this arm
        // the cast hid its own operand from this pass: `v(1)` kept its Fortran
        // 1-based subscript and read `v(2)`.
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Cast { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner) => {
            rewrite_array_subscripts_in_expr(inner, arrays, char_vars, callables, array_fields)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_array_subscripts_in_expr(cond, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(then, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(else_, arrays, char_vars, callables, array_fields);
        }
        ExprKind::Member { object, .. } => {
            rewrite_array_subscripts_in_expr(object, arrays, char_vars, callables, array_fields)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_array_subscripts_in_expr(object, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(index, arrays, char_vars, callables, array_fields);

            if let ExprKind::Ident(var_name) = &object.kind {
                if char_vars.contains(&var_name.to_ascii_lowercase()) {
                    if let ExprKind::Slice { lower, upper, .. } = &index.kind {
                        let start = lower.as_deref().map_or_else(
                            || Expression::int(0),
                            |lower| {
                                Expression::new(ExprKind::Binary {
                                    left: Box::new(lower.clone()),
                                    op: BinOp::Sub,
                                    right: Box::new(Expression::int(1)),
                                })
                            },
                        );
                        let end = upper.as_deref().cloned();
                        *expr = build_fortran_str_slice(Expression::ident(var_name), start, end);
                        return;
                    }
                    if matches!(&index.kind, ExprKind::Range { .. }) {
                        return;
                    }
                }
            }

            *index = Box::new(normalize_array_index_operand(
                index.as_ref().clone(),
                FORTRAN_ARRAY_INDEXING,
            ));
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_array_subscripts_in_expr(lower, arrays, char_vars, callables, array_fields);
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_array_subscripts_in_expr(upper, arrays, char_vars, callables, array_fields);
            }
            if let Some(step) = step.as_mut() {
                rewrite_array_subscripts_in_expr(step, arrays, char_vars, callables, array_fields);
            }
        }
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            rewrite_array_subscripts_in_expr(callee, arrays, char_vars, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_array_subscripts_in_expr(
                    &mut arg.value,
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
            if !args.is_empty()
                && !*optional
                && !is_known_fortran_callable(callee, callables)
                && (is_known_fortran_array(callee, arrays, array_fields)
                    || matches!(&callee.kind, ExprKind::Index { .. }))
            {
                expr.kind = build_fortran_index_chain(callee.as_ref().clone(), args);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_array_subscripts_in_expr(class, arrays, char_vars, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_array_subscripts_in_expr(
                    &mut arg.value,
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_array_subscripts_in_expr(target, arrays, char_vars, callables, array_fields);
            rewrite_array_subscripts_in_expr(value, arrays, char_vars, callables, array_fields);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    rewrite_array_subscripts_in_expr(
                        key,
                        arrays,
                        char_vars,
                        callables,
                        array_fields,
                    );
                }
                rewrite_array_subscripts_in_expr(
                    &mut item.value,
                    arrays,
                    char_vars,
                    callables,
                    array_fields,
                );
            }
        }
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args {
                rewrite_array_subscripts_in_expr(arg, arrays, char_vars, callables, array_fields);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_array_subscripts_in_expr(item, arrays, char_vars, callables, array_fields);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_array_subscripts_in_expr(
                            key,
                            arrays,
                            char_vars,
                            callables,
                            array_fields,
                        );
                        rewrite_array_subscripts_in_expr(
                            value,
                            arrays,
                            char_vars,
                            callables,
                            array_fields,
                        );
                    }
                    ObjectProperty::Spread(expr) => rewrite_array_subscripts_in_expr(
                        expr,
                        arrays,
                        char_vars,
                        callables,
                        array_fields,
                    ),
                    ObjectProperty::Computed { key, value } => {
                        rewrite_array_subscripts_in_expr(
                            key,
                            arrays,
                            char_vars,
                            callables,
                            array_fields,
                        );
                        rewrite_array_subscripts_in_expr(
                            value,
                            arrays,
                            char_vars,
                            callables,
                            array_fields,
                        );
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_array_subscripts_in_expr(
                            expr,
                            arrays,
                            char_vars,
                            callables,
                            array_fields,
                        )
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn build_fortran_index_chain(object: Expression, args: &[Argument]) -> ExprKind {
    build_fortran_index_chain_expr(object, args, 0).kind
}

fn build_fortran_index_chain_expr(
    object: Expression,
    args: &[Argument],
    depth: usize,
) -> Expression {
    let Some((first, rest)) = args.split_first() else {
        return object;
    };

    let indexed = Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(normalize_array_index_operand(
            first.value.clone(),
            FORTRAN_ARRAY_INDEXING,
        )),
        null_safe: false,
    });

    if rest.is_empty() {
        return indexed;
    }
    if matches!(first.value.kind, ExprKind::Slice { .. }) {
        let item_name = format!("__fortran_section_item_{depth}");
        let index_name = format!("__fortran_section_index_{depth}");
        return build_fortran_array_map(
            indexed,
            build_fortran_index_chain_expr(Expression::ident(&item_name), rest, depth + 1),
            false,
            &item_name,
            &index_name,
        );
    }

    build_fortran_index_chain_expr(indexed, rest, depth + 1)
}

fn is_known_fortran_array(
    expr: &Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => arrays.contains(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_fields.contains(&field.to_ascii_lowercase()),
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => matches!(
                name.to_ascii_lowercase().as_str(),
                "str_split" | "str_getcsv" | "array"
            ),
            _ => false,
        },
        ExprKind::Index { .. } => true,
        _ => false,
    }
}

fn is_known_fortran_callable(expr: &Expression, callables: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => callables.contains(&name.to_ascii_lowercase()),
        _ => false,
    }
}

fn is_fortran_callable_type_hint(type_hint: &str) -> bool {
    let lower = type_hint.trim().to_ascii_lowercase();
    lower.starts_with("procedure(") || lower == "procedure"
}

fn fortran_callable_signature_name(type_hint: &str) -> Option<String> {
    type_hint
        .trim()
        .strip_prefix("procedure(")
        .and_then(|rest| rest.strip_suffix(')'))
        .map(|name| name.trim().to_ascii_lowercase())
}

fn collect_fortran_array_procedure_params(
    params: &[Param],
    array_functions: &HashSet<String>,
) -> HashSet<String> {
    params
        .iter()
        .filter_map(|param| {
            if param.type_hint.is_none() && matches!(param.pass_by, PassBy::Value) {
                return Some(param.name.to_ascii_lowercase());
            }

            param
                .type_hint
                .as_deref()
                .and_then(fortran_callable_signature_name)
                .filter(|signature_name| array_functions.contains(signature_name))
                .map(|_| param.name.to_ascii_lowercase())
        })
        .collect()
}

fn is_fortran_string_type_hint(type_hint: &str) -> bool {
    let lower = type_hint.trim().to_ascii_lowercase();
    lower == "character" || lower.starts_with("character(") || lower.starts_with("character*")
}

fn is_array_initializer_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Array(_) | ExprKind::ArrayTransform { .. } => true,
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array"))
        }
        _ => false,
    }
}

fn lower_fortran_body_intrinsics(params: &[Param], body: &mut Vec<Statement>) {
    let mut type_env = HashMap::new();
    for param in params {
        if let Some(type_hint) = &param.type_hint {
            type_env.insert(
                param.name.to_ascii_lowercase(),
                type_hint.spelling().to_string(),
            );
        }
    }
    lower_fortran_body_intrinsics_with_env(body, &mut type_env);
}

fn lower_body_intrinsic_statement(
    statement: &Statement,
    type_env: &HashMap<String, String>,
) -> Option<Statement> {
    if let StmtKind::Expr(expr) = &statement.kind {
        if let Some(stmt) = lower_fortran_random_statement_with_env(expr, type_env) {
            return Some(stmt);
        }
        return lower_intrinsic_statement(expr);
    }
    None
}

fn lower_fortran_random_statement_with_env(
    expr: &Expression,
    type_env: &HashMap<String, String>,
) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name.eq_ignore_ascii_case("random_seed") {
        return Some(lower_fortran_random_seed_statement(args));
    }
    if !name.eq_ignore_ascii_case("random_number") {
        return None;
    }
    let target = args
        .iter()
        .find(|arg| {
            arg.name
                .as_deref()
                .is_none_or(|name| name.eq_ignore_ascii_case("harvest"))
        })?
        .value
        .clone();
    let assign_target = fortran_random_assignment_target(target.clone());
    let value = if fortran_random_target_is_array(&assign_target, type_env) {
        fortran_random_array_fill_expr(assign_target.clone())
    } else {
        Expression::float(0.5)
    };
    Some(Statement::new(StmtKind::Assign {
        targets: vec![assign_target],
        value,
        by_ref: false,
    }))
}

fn lower_fortran_body_intrinsics_with_env(
    body: &mut [Statement],
    type_env: &mut HashMap<String, String>,
) {
    for statement in body.iter_mut() {
        if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                let Some(type_hint) = &declaration.type_hint else {
                    continue;
                };
                type_env.insert(
                    name.to_ascii_lowercase(),
                    fortran_array_type_hint(type_hint, declaration.array_bounds.as_deref()),
                );
            }
        }

        lower_fortran_type_inquiry_in_statement(statement, type_env);

        if let Some(lowered) = lower_body_intrinsic_statement(statement, type_env) {
            *statement = lowered;
        }
        if let Some(lowered) = lower_fortran_transfer_assignment_statement(statement, type_env) {
            *statement = lowered;
        }
        lower_fortran_transfer_markers_in_statement(statement, type_env);

        match &mut statement.kind {
            StmtKind::VarDecl { .. } => {}
            StmtKind::FunctionDecl { params, body, .. } => {
                let mut nested_env = type_env.clone();
                for param in params {
                    if let Some(type_hint) = &param.type_hint {
                        nested_env.insert(
                            param.name.to_ascii_lowercase(),
                            type_hint.clone().to_string(),
                        );
                    }
                }
                lower_fortran_body_intrinsics_with_env(body, &mut nested_env);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                let mut nested_env = type_env.clone();
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            lower_fortran_body_intrinsics_with_env(
                                std::slice::from_mut(stmt.as_mut()),
                                &mut nested_env,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts) | StmtKind::DoWhile { body: stmts, .. } => {
                let mut nested_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut nested_env);
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_body_intrinsics_with_env(
                    std::slice::from_mut(body.as_mut()),
                    type_env,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(then_body, &mut then_env);
                for (_, elif_body) in elifs {
                    let mut elif_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(elif_body, &mut elif_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_env = type_env.clone();
                if let Some(init) = init {
                    lower_fortran_body_intrinsics_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_env,
                    );
                }
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
            }
            StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(&mut case.body, &mut case_env);
                }
                if let Some(default) = default {
                    let mut default_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(default, &mut default_env);
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(try_body, &mut try_env);
                for catch in catches {
                    let mut catch_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(&mut catch.body, &mut catch_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
                if let Some(finally) = finally {
                    let mut finally_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(finally, &mut finally_env);
                }
            }
            StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut nested_env);
            }
            _ => {}
        }
    }
}

fn lower_fortran_transfer_assignment_statement(
    statement: &Statement,
    type_env: &HashMap<String, String>,
) -> Option<Statement> {
    let StmtKind::Assign {
        targets,
        value,
        by_ref,
    } = &statement.kind
    else {
        return None;
    };
    if targets.len() != 1 {
        return None;
    }
    let ExprKind::Call { callee, args, .. } = &value.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !is_fortran_transfer_marker_name(name) {
        return None;
    }
    let positional = args
        .iter()
        .filter(|arg| arg.name.is_none())
        .map(|arg| arg.value.clone())
        .collect::<Vec<_>>();
    if positional.len() < 2 {
        return None;
    }
    let target = targets[0].clone();
    let target_hint = fortran_type_hint_for_expr(&target, type_env);
    Some(Statement::new(StmtKind::Assign {
        targets: vec![target],
        value: build_fortran_transfer_expr_with_hint(
            positional[0].clone(),
            positional[1].clone(),
            positional.get(2).cloned(),
            target_hint.as_deref(),
            fortran_type_hint_for_expr(&positional[0], type_env).as_deref(),
        ),
        by_ref: *by_ref,
    }))
}

fn is_fortran_transfer_marker_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("transfer") || name.eq_ignore_ascii_case("__fortran_transfer")
}

fn lower_fortran_transfer_markers_in_statement(
    statement: &mut Statement,
    type_env: &HashMap<String, String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            lower_fortran_transfer_markers_in_expr(expr, type_env)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations {
                if let Some(init) = &mut declaration.init {
                    lower_fortran_transfer_markers_in_expr(init, type_env);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                lower_fortran_transfer_markers_in_expr(target, type_env);
            }
            lower_fortran_transfer_markers_in_expr(value, type_env);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            lower_fortran_transfer_markers_in_expr(target, type_env);
            lower_fortran_transfer_markers_in_expr(value, type_env);
        }
        StmtKind::If {
            cond,
            elifs,
            else_body: _,
            ..
        } => {
            lower_fortran_transfer_markers_in_expr(cond, type_env);
            for (elif_cond, _) in elifs {
                lower_fortran_transfer_markers_in_expr(elif_cond, type_env);
            }
        }
        StmtKind::While { cond, .. } | StmtKind::DoWhile { cond, .. } => {
            lower_fortran_transfer_markers_in_expr(cond, type_env)
        }
        StmtKind::For { init, cond, update, .. } => {
            if let Some(init) = init {
                lower_fortran_transfer_markers_in_statement(init, type_env);
            }
            if let Some(cond) = cond {
                lower_fortran_transfer_markers_in_expr(cond, type_env);
            }
            if let Some(update) = update {
                lower_fortran_transfer_markers_in_expr(update, type_env);
            }
        }
        StmtKind::ForIn { iter, .. } => lower_fortran_transfer_markers_in_expr(iter, type_env),
        StmtKind::Switch { expr, cases, .. } => {
            lower_fortran_transfer_markers_in_expr(expr, type_env);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(value) => {
                            lower_fortran_transfer_markers_in_expr(value, type_env)
                        }
                        CaseCondition::Range { from, to } => {
                            lower_fortran_transfer_markers_in_expr(from, type_env);
                            lower_fortran_transfer_markers_in_expr(to, type_env);
                        }
                        // An open-ended range carries its one written bound
                        // here — skipping it would leave a marker unlowered.
                        CaseCondition::Comparison { expr, .. } => {
                            lower_fortran_transfer_markers_in_expr(expr, type_env)
                        }
                    }
                }
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                lower_fortran_transfer_markers_in_expr(expr, type_env);
            }
            if let Some(cause) = cause {
                lower_fortran_transfer_markers_in_expr(cause, type_env);
            }
        }
        StmtKind::Assert { test, msg } => {
            lower_fortran_transfer_markers_in_expr(test, type_env);
            if let Some(msg) = msg {
                lower_fortran_transfer_markers_in_expr(msg, type_env);
            }
        }
        // `print *, transfer(n, 0.0)` carries the marker in an item list; without
        // these arms the marker survives into emitted code as a call to undefined.
        StmtKind::Echo(items) => {
            for item in items {
                lower_fortran_transfer_markers_in_expr(item, type_env);
            }
        }
        StmtKind::PrintFile {
            file_number, items, ..
        }
        | StmtKind::WriteFile {
            file_number, items, ..
        } => {
            lower_fortran_transfer_markers_in_expr(file_number, type_env);
            for item in items {
                lower_fortran_transfer_markers_in_expr(item, type_env);
            }
        }
        _ => {}
    }
}

fn lower_fortran_transfer_markers_in_expr(
    expr: &mut Expression,
    type_env: &HashMap<String, String>,
) {
    match &mut expr.kind {
        // `MERGE`/`PACK`/`UNPACK`/`RESHAPE` are ArrayTransform NODES, and
        // their arguments are ORDINARY expressions. Without an arm here the
        // pass walks straight past them — which is how `nearest(...)` inside
        // a `merge(...)` stopped being folded the moment MERGE became a node
        // instead of a call.
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args.iter_mut() {
                lower_fortran_transfer_markers_in_expr(arg, type_env);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            lower_fortran_transfer_markers_in_expr(callee, type_env);
            for arg in args.iter_mut() {
                lower_fortran_transfer_markers_in_expr(&mut arg.value, type_env);
            }
            let ExprKind::Ident(name) = &callee.kind else {
                return;
            };
            if !is_fortran_transfer_marker_name(name) {
                return;
            }
            let positional = args
                .iter()
                .filter(|arg| arg.name.is_none())
                .map(|arg| arg.value.clone())
                .collect::<Vec<_>>();
            if positional.len() < 2 {
                return;
            }
            *expr = build_fortran_transfer_expr_with_hint(
                positional[0].clone(),
                positional[1].clone(),
                positional.get(2).cloned(),
                None,
                fortran_type_hint_for_expr(&positional[0], type_env).as_deref(),
            );
        }
        ExprKind::Binary { left, right, .. } => {
            lower_fortran_transfer_markers_in_expr(left, type_env);
            lower_fortran_transfer_markers_in_expr(right, type_env);
        }
        ExprKind::Unary { expr: inner, .. } => {
            lower_fortran_transfer_markers_in_expr(inner, type_env);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            lower_fortran_transfer_markers_in_expr(cond, type_env);
            lower_fortran_transfer_markers_in_expr(then, type_env);
            lower_fortran_transfer_markers_in_expr(else_, type_env);
        }
        ExprKind::Member { object, .. } => lower_fortran_transfer_markers_in_expr(object, type_env),
        ExprKind::Index { object, index, .. } => {
            lower_fortran_transfer_markers_in_expr(object, type_env);
            lower_fortran_transfer_markers_in_expr(index, type_env);
        }
        // `print *, "x", transfer(s, 0), "]"` — two or more items become an
        // Interpolation, and without this arm the marker sails through into
        // emitted code as a call to undefined. A ONE-item print skips
        // Interpolation entirely, which is what made this look like it worked.
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(part) = part {
                    lower_fortran_transfer_markers_in_expr(part, type_env);
                }
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                lower_fortran_transfer_markers_in_expr(&mut item.value, type_env);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value } => {
                        lower_fortran_transfer_markers_in_expr(key, type_env);
                        lower_fortran_transfer_markers_in_expr(value, type_env);
                    }
                    ObjectProperty::Spread(value) => {
                        lower_fortran_transfer_markers_in_expr(value, type_env)
                    }
                    ObjectProperty::Computed { key, value } => {
                        lower_fortran_transfer_markers_in_expr(key, type_env);
                        lower_fortran_transfer_markers_in_expr(value, type_env);
                    }
                    ObjectProperty::Method { .. } | ObjectProperty::Accessor { .. } => {}
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn fortran_type_hint_for_expr(
    expr: &Expression,
    type_env: &HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => type_env.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Member { field, .. } => type_env.get(&field.to_ascii_lowercase()).cloned(),
        _ => None,
    }
}

fn fortran_type_hint_is_array(type_hint: Option<&str>) -> bool {
    type_hint.is_some_and(|hint| hint.contains("()"))
}

fn is_fortran_complex_type_hint(type_hint: &str) -> bool {
    let lower = strip_fortran_type_hint_array_rank(type_hint)
        .trim()
        .to_ascii_lowercase();
    lower.starts_with("complex") || lower == "double complex"
}

fn is_fortran_scalar_complex_type_hint(type_hint: &str) -> bool {
    is_fortran_complex_type_hint(type_hint) && fortran_type_hint_array_rank(type_hint) == 0
}

fn is_fortran_array_complex_type_hint(type_hint: &str) -> bool {
    is_fortran_complex_type_hint(type_hint) && fortran_type_hint_array_rank(type_hint) > 0
}

fn lower_fortran_complex_expressions(body: &mut [Statement]) {
    let mut type_env = HashMap::new();
    collect_fortran_complex_type_env(body, &mut type_env);
    lower_fortran_complex_expressions_with_env(body, &mut type_env);
}

fn collect_fortran_complex_type_env(body: &[Statement], type_env: &mut HashMap<String, String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(type_hint) = &declaration.type_hint else {
                        continue;
                    };
                    type_env.insert(
                        name.to_ascii_lowercase(),
                        fortran_array_type_hint(type_hint, declaration.array_bounds.as_deref()),
                    );
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                for param in params {
                    if let Some(type_hint) = &param.type_hint {
                        type_env.insert(
                            param.name.to_ascii_lowercase(),
                            type_hint.clone().to_string(),
                        );
                    }
                }
                collect_fortran_complex_type_env(body, type_env);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            collect_fortran_complex_type_env(
                                std::slice::from_ref(stmt.as_ref()),
                                type_env,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                collect_fortran_complex_type_env(stmts, type_env);
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                collect_fortran_complex_type_env(std::slice::from_ref(body.as_ref()), type_env);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_fortran_complex_type_env(then_body, type_env);
                for (_, elif_body) in elifs {
                    collect_fortran_complex_type_env(elif_body, type_env);
                }
                if let Some(else_body) = else_body {
                    collect_fortran_complex_type_env(else_body, type_env);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                collect_fortran_complex_type_env(stmts, type_env);
                if let Some(else_body) = else_body {
                    collect_fortran_complex_type_env(else_body, type_env);
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                if let Some(init) = init {
                    collect_fortran_complex_type_env(std::slice::from_ref(init.as_ref()), type_env);
                }
                collect_fortran_complex_type_env(stmts, type_env);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    collect_fortran_complex_type_env(&case.body, type_env);
                }
                if let Some(default) = default {
                    collect_fortran_complex_type_env(default, type_env);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                collect_fortran_complex_type_env(body, type_env);
                for catch in catches {
                    collect_fortran_complex_type_env(&catch.body, type_env);
                }
                if let Some(else_body) = else_body {
                    collect_fortran_complex_type_env(else_body, type_env);
                }
                if let Some(finally) = finally {
                    collect_fortran_complex_type_env(finally, type_env);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_complex_expressions_with_env(
    body: &mut [Statement],
    type_env: &mut HashMap<String, String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_complex_expressions_in_statement(statement, type_env);

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(type_hint) = &declaration.type_hint else {
                        continue;
                    };
                    type_env.insert(
                        name.to_ascii_lowercase(),
                        fortran_array_type_hint(type_hint, declaration.array_bounds.as_deref()),
                    );
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                let mut nested_env = type_env.clone();
                for param in params {
                    if let Some(type_hint) = &param.type_hint {
                        nested_env.insert(
                            param.name.to_ascii_lowercase(),
                            type_hint.clone().to_string(),
                        );
                    }
                }
                lower_fortran_complex_expressions_with_env(body, &mut nested_env);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                let mut nested_env = type_env.clone();
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            lower_fortran_complex_expressions_with_env(
                                std::slice::from_mut(stmt.as_mut()),
                                &mut nested_env,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested_env = type_env.clone();
                lower_fortran_complex_expressions_with_env(stmts, &mut nested_env);
            }
            // A named construct wraps its statement in `Labeled`; transparent.
            StmtKind::Labeled { body, .. } => {
                lower_fortran_complex_expressions_with_env(
                    std::slice::from_mut(body.as_mut()),
                    type_env,
                );
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut then_env = type_env.clone();
                lower_fortran_complex_expressions_with_env(then_body, &mut then_env);
                for (_, elif_body) in elifs {
                    let mut elif_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(elif_body, &mut elif_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::While {
                body: stmts,
                else_body,
                ..
            }
            | StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                let mut loop_env = type_env.clone();
                lower_fortran_complex_expressions_with_env(stmts, &mut loop_env);
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::For {
                init, body: stmts, ..
            } => {
                let mut loop_env = type_env.clone();
                if let Some(init) = init {
                    lower_fortran_complex_expressions_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_env,
                    );
                }
                lower_fortran_complex_expressions_with_env(stmts, &mut loop_env);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(&mut case.body, &mut case_env);
                }
                if let Some(default) = default {
                    let mut default_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(default, &mut default_env);
                }
            }
            StmtKind::Try {
                body: try_body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_env = type_env.clone();
                lower_fortran_complex_expressions_with_env(try_body, &mut try_env);
                for catch in catches {
                    let mut catch_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(&mut catch.body, &mut catch_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(else_body, &mut else_env);
                }
                if let Some(finally) = finally {
                    let mut finally_env = type_env.clone();
                    lower_fortran_complex_expressions_with_env(finally, &mut finally_env);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_complex_expressions_in_statement(
    statement: &mut Statement,
    type_env: &HashMap<String, String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_fortran_complex_expressions_in_expr(expr, type_env);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_fortran_complex_expressions_in_expr(expr, type_env);
            }
            if let Some(cause) = cause {
                rewrite_fortran_complex_expressions_in_expr(cause, type_env);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations {
                if let Some(init) = &mut declaration.init {
                    rewrite_fortran_complex_expressions_in_expr(init, type_env);
                }
                if let Some(bounds) = &mut declaration.array_bounds {
                    for bound in bounds {
                        rewrite_fortran_complex_expressions_in_expr(bound, type_env);
                    }
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            let target_is_complex_array = targets
                .iter()
                .any(|target| expr_is_fortran_complex_array_target(target, type_env))
                || matches!(&value.kind,
                    ExprKind::Call { callee, .. }
                        if matches!(&callee.kind,
                            ExprKind::Member { object, field, .. }
                                if field.eq_ignore_ascii_case("map")
                                    && expr_is_fortran_complex_array(object, type_env)
                        )
                );
            for target in targets {
                rewrite_fortran_complex_expressions_in_expr(target, type_env);
            }
            rewrite_fortran_complex_expressions_in_expr(value, type_env);
            if target_is_complex_array {
                force_fortran_complex_array_value(value, type_env);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_fortran_complex_expressions_in_expr(target, type_env);
            rewrite_fortran_complex_expressions_in_expr(value, type_env);
        }
        StmtKind::If { cond, .. }
        | StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. } => {
            rewrite_fortran_complex_expressions_in_expr(cond, type_env);
        }
        StmtKind::For {
            init, cond, update, ..
        } => {
            if let Some(init) = init {
                rewrite_fortran_complex_expressions_in_statement(init, type_env);
            }
            if let Some(cond) = cond {
                rewrite_fortran_complex_expressions_in_expr(cond, type_env);
            }
            if let Some(update) = update {
                rewrite_fortran_complex_expressions_in_expr(update, type_env);
            }
        }
        StmtKind::ForIn { iter, .. } => {
            rewrite_fortran_complex_expressions_in_expr(iter, type_env);
        }
        StmtKind::Switch { expr, cases, .. } => {
            rewrite_fortran_complex_expressions_in_expr(expr, type_env);
            for case in cases {
                for condition in &mut case.conditions {
                    rewrite_fortran_complex_expressions_in_case_condition(condition, type_env);
                }
            }
        }
        StmtKind::Echo(items) => {
            for item in items {
                rewrite_fortran_complex_expressions_in_expr(item, type_env);
            }
        }
        StmtKind::Assert { test, msg } => {
            rewrite_fortran_complex_expressions_in_expr(test, type_env);
            if let Some(msg) = msg {
                rewrite_fortran_complex_expressions_in_expr(msg, type_env);
            }
        }
        _ => {}
    }
}

fn rewrite_fortran_complex_expressions_in_case_condition(
    condition: &mut CaseCondition,
    type_env: &HashMap<String, String>,
) {
    match condition {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            rewrite_fortran_complex_expressions_in_expr(expr, type_env);
        }
        CaseCondition::Range { from, to } => {
            rewrite_fortran_complex_expressions_in_expr(from, type_env);
            rewrite_fortran_complex_expressions_in_expr(to, type_env);
        }
    }
}

fn rewrite_fortran_complex_expressions_in_expr(
    expr: &mut Expression,
    type_env: &HashMap<String, String>,
) {
    match &mut expr.kind {
        // `real(z)` and `int(z)` are folded to a numeric cast while the
        // expression is being walked, before any declaration is known — the same
        // way `a + b` arrives here as a plain `Binary`. On a complex operand a
        // numeric cast is not what either one means: both take the REAL PART,
        // and `int` truncates it afterwards.
        ExprKind::Cast { expr: inner, type_name } => {
            rewrite_fortran_complex_expressions_in_expr(inner, type_env);
            if !expr_is_fortran_complex_scalar(inner, type_env) {
                return;
            }
            let real_part = fortran_complex_real_part(inner);
            *expr = match type_name.as_str() {
                "number" => real_part,
                _ => Expression::new(ExprKind::Cast {
                    expr: Box::new(real_part),
                    type_name: type_name.clone(),
                }),
            };
        }
        ExprKind::Binary { op, left, right } => {
            rewrite_fortran_complex_expressions_in_expr(left, type_env);
            rewrite_fortran_complex_expressions_in_expr(right, type_env);
            if matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div)
                && (expr_is_fortran_complex_scalar(left, type_env)
                    || expr_is_fortran_complex_scalar(right, type_env))
            {
                *expr = lower_fortran_complex_binary_expr(*op, left, right, type_env);
            }
        }
        ExprKind::Unary { op, expr: inner } => {
            rewrite_fortran_complex_expressions_in_expr(inner, type_env);
            if matches!(op, UnaryOp::Neg) && expr_is_fortran_complex_scalar(inner, type_env) {
                *expr = build_fortran_complex_expr(
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Neg,
                        expr: Box::new(fortran_complex_real_part(inner)),
                    }),
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Neg,
                        expr: Box::new(fortran_complex_imag_part(inner)),
                    }),
                );
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_fortran_complex_expressions_in_expr(object, type_env);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_fortran_complex_expressions_in_expr(object, type_env);
            rewrite_fortran_complex_expressions_in_expr(index, type_env);
        }
        // `MERGE`/`PACK`/`UNPACK`/`RESHAPE` are ArrayTransform NODES, and
        // their arguments are ORDINARY expressions. Without an arm here the
        // pass walks straight past them — which is how `nearest(...)` inside
        // a `merge(...)` stopped being folded the moment MERGE became a node
        // instead of a call.
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args.iter_mut() {
                rewrite_fortran_complex_expressions_in_expr(arg, type_env);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_fortran_complex_expressions_in_expr(callee, type_env);
            let map_item_type_hint = match &callee.kind {
                ExprKind::Member { object, field, .. }
                    if field.eq_ignore_ascii_case("map")
                        && expr_is_fortran_complex_array(object, type_env) =>
                {
                    Some("complex".to_string())
                }
                _ => None,
            };
            for (index, arg) in args.iter_mut().enumerate() {
                if index == 0 {
                    if let Some(item_type_hint) = &map_item_type_hint {
                        if let ExprKind::Lambda { params, body, .. } = &mut arg.value.kind {
                            let mut nested_env = type_env.clone();
                            if let Some(first_param) = params.first_mut() {
                                if first_param.type_hint.is_none() {
                                    first_param.type_hint = Some(item_type_hint.clone().into());
                                }
                                nested_env.insert(
                                    first_param.name.to_ascii_lowercase(),
                                    item_type_hint.clone(),
                                );
                            }
                            rewrite_fortran_complex_lambda_body(body, &nested_env);
                            continue;
                        }
                    }
                }
                rewrite_fortran_complex_expressions_in_expr(&mut arg.value, type_env);
            }
            if let Some(lowered) = lower_fortran_complex_call_expr(callee, args, type_env) {
                *expr = lowered;
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_fortran_complex_expressions_in_expr(target, type_env);
            rewrite_fortran_complex_expressions_in_expr(value, type_env);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut nested_env = type_env.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    nested_env.insert(
                        param.name.to_ascii_lowercase(),
                        type_hint.clone().to_string(),
                    );
                }
            }
            rewrite_fortran_complex_lambda_body(body, &nested_env);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    rewrite_fortran_complex_expressions_in_expr(key, type_env);
                }
                rewrite_fortran_complex_expressions_in_expr(&mut item.value, type_env);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_fortran_complex_expressions_in_expr(key, type_env);
                        rewrite_fortran_complex_expressions_in_expr(value, type_env);
                    }
                    ObjectProperty::Spread(inner) => {
                        rewrite_fortran_complex_expressions_in_expr(inner, type_env);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_fortran_complex_expressions_in_expr(lower, type_env);
            }
            if let Some(upper) = upper {
                rewrite_fortran_complex_expressions_in_expr(upper, type_env);
            }
            if let Some(step) = step {
                rewrite_fortran_complex_expressions_in_expr(step, type_env);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_fortran_complex_expressions_in_expr(cond, type_env);
            rewrite_fortran_complex_expressions_in_expr(then, type_env);
            rewrite_fortran_complex_expressions_in_expr(else_, type_env);
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_fortran_complex_expressions_in_expr(left, type_env);
            rewrite_fortran_complex_expressions_in_expr(right, type_env);
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(inner) | InterpolPart::Formatted(inner, _) => {
                        rewrite_fortran_complex_expressions_in_expr(inner, type_env);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn rewrite_fortran_complex_lambda_body(body: &mut LambdaBody, type_env: &HashMap<String, String>) {
    match body {
        LambdaBody::Expr(expr) => rewrite_fortran_complex_expressions_in_expr(expr, type_env),
        LambdaBody::Block(stmts) => {
            let mut nested_env = type_env.clone();
            lower_fortran_complex_expressions_with_env(stmts, &mut nested_env);
        }
    }
}

fn expr_is_fortran_complex_array_target(
    expr: &Expression,
    type_env: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .is_some_and(|type_hint| is_fortran_array_complex_type_hint(type_hint)),
        ExprKind::Member { object, .. } => expr_is_fortran_complex_array_target(object, type_env),
        ExprKind::Index { object, index, .. } => {
            matches!(index.kind, ExprKind::Slice { .. })
                && expr_is_fortran_complex_array_target(object, type_env)
        }
        _ => false,
    }
}

fn force_fortran_complex_array_value(expr: &mut Expression, type_env: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            force_fortran_complex_array_value(left, type_env);
            force_fortran_complex_array_value(right, type_env);
        }
        ExprKind::Unary { expr: inner, .. } | ExprKind::Member { object: inner, .. } => {
            force_fortran_complex_array_value(inner, type_env);
        }
        ExprKind::Index { object, index, .. } => {
            force_fortran_complex_array_value(object, type_env);
            force_fortran_complex_array_value(index, type_env);
        }
        ExprKind::Call { callee, args, .. } => {
            force_fortran_complex_array_value(callee, type_env);
            let is_map = matches!(&callee.kind, ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("map"));
            for (index, arg) in args.iter_mut().enumerate() {
                if is_map && index == 0 {
                    if let ExprKind::Lambda { params, body, .. } = &mut arg.value.kind {
                        let mut nested_env = type_env.clone();
                        if let Some(first_param) = params.first_mut() {
                            if first_param.type_hint.is_none() {
                                first_param.type_hint = Some("complex".to_string().into());
                            }
                            nested_env.insert(
                                first_param.name.to_ascii_lowercase(),
                                "complex".to_string(),
                            );
                        }
                        match body {
                            LambdaBody::Expr(expr) => {
                                if let ExprKind::Binary { op, left, right } = &expr.kind {
                                    if matches!(
                                        op,
                                        BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div
                                    ) {
                                        **expr = lower_fortran_complex_binary_expr(
                                            *op,
                                            left,
                                            right,
                                            &nested_env,
                                        );
                                        continue;
                                    }
                                }
                                rewrite_fortran_complex_expressions_in_expr(expr, &nested_env);
                            }
                            LambdaBody::Block(stmts) => {
                                let mut block_env = nested_env.clone();
                                lower_fortran_complex_expressions_with_env(stmts, &mut block_env);
                            }
                        }
                        continue;
                    }
                }
                force_fortran_complex_array_value(&mut arg.value, type_env);
            }
        }
        ExprKind::Assign { target, value } => {
            force_fortran_complex_array_value(target, type_env);
            force_fortran_complex_array_value(value, type_env);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut nested_env = type_env.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    nested_env.insert(
                        param.name.to_ascii_lowercase(),
                        type_hint.clone().to_string(),
                    );
                }
            }
            rewrite_fortran_complex_lambda_body(body, &nested_env);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    force_fortran_complex_array_value(key, type_env);
                }
                force_fortran_complex_array_value(&mut item.value, type_env);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        force_fortran_complex_array_value(key, type_env);
                        force_fortran_complex_array_value(value, type_env);
                    }
                    ObjectProperty::Spread(inner) => {
                        force_fortran_complex_array_value(inner, type_env);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                force_fortran_complex_array_value(lower, type_env);
            }
            if let Some(upper) = upper {
                force_fortran_complex_array_value(upper, type_env);
            }
            if let Some(step) = step {
                force_fortran_complex_array_value(step, type_env);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            force_fortran_complex_array_value(cond, type_env);
            force_fortran_complex_array_value(then, type_env);
            force_fortran_complex_array_value(else_, type_env);
        }
        ExprKind::NullCoalesce { left, right } => {
            force_fortran_complex_array_value(left, type_env);
            force_fortran_complex_array_value(right, type_env);
        }
        _ => {}
    }
}

fn expr_is_fortran_complex_scalar(expr: &Expression, type_env: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .is_some_and(|type_hint| is_fortran_scalar_complex_type_hint(type_hint)),
        ExprKind::Index { object, index, .. } => {
            !matches!(index.kind, ExprKind::Slice { .. })
                && expr_is_fortran_complex_array_base(object, type_env)
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                let lowered = name.to_ascii_lowercase();
                lowered == "cmplx"
                    || lowered == "conjg"
                    || type_env
                        .get(&lowered)
                        .is_some_and(|type_hint| is_fortran_scalar_complex_type_hint(type_hint))
                    || (type_env
                        .get(&lowered)
                        .is_some_and(|type_hint| is_fortran_array_complex_type_hint(type_hint))
                        && !args
                            .iter()
                            .any(|arg| matches!(arg.value.kind, ExprKind::Slice { .. })))
            }
            _ => false,
        },
        ExprKind::Object(props) => fortran_complex_object_fields(props),
        ExprKind::Unary { expr, .. } => expr_is_fortran_complex_scalar(expr, type_env),
        ExprKind::Binary { left, right, .. } => {
            expr_is_fortran_complex_scalar(left, type_env)
                || expr_is_fortran_complex_scalar(right, type_env)
        }
        _ => false,
    }
}

fn expr_is_fortran_complex_array(expr: &Expression, type_env: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .is_some_and(|type_hint| is_fortran_array_complex_type_hint(type_hint)),
        ExprKind::Index { object, index, .. } => {
            matches!(index.kind, ExprKind::Slice { .. })
                && expr_is_fortran_complex_array_base(object, type_env)
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                type_env
                    .get(&name.to_ascii_lowercase())
                    .is_some_and(|type_hint| is_fortran_array_complex_type_hint(type_hint))
                    && args
                        .iter()
                        .any(|arg| matches!(arg.value.kind, ExprKind::Slice { .. }))
            }
            _ => false,
        },
        _ => false,
    }
}

fn expr_is_fortran_complex_array_base(
    expr: &Expression,
    type_env: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .is_some_and(|type_hint| is_fortran_array_complex_type_hint(type_hint)),
        _ => expr_is_fortran_complex_array(expr, type_env),
    }
}

fn fortran_complex_object_fields(props: &[ObjectProperty]) -> bool {
    let mut has_re = false;
    let mut has_im = false;
    for prop in props {
        let ObjectProperty::KeyValue { key, .. } = prop else {
            continue;
        };
        let ExprKind::Lit(Literal::Str(name)) = &key.kind else {
            continue;
        };
        if name.eq_ignore_ascii_case("real") {
            has_re = true;
        } else if name.eq_ignore_ascii_case("imag") {
            has_im = true;
        }
    }
    has_re && has_im
}

fn lower_fortran_complex_call_expr(
    callee: &Expression,
    args: &[Argument],
    type_env: &HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let lowered = name.to_ascii_lowercase();
    let positional_args = args
        .iter()
        .filter(|arg| {
            arg.name
                .as_deref()
                .is_none_or(|name| !name.eq_ignore_ascii_case("kind"))
        })
        .map(|arg| arg.value.clone())
        .collect::<Vec<_>>();
    match lowered.as_str() {
        "cmplx" => Some(build_fortran_complex_expr(
            positional_args
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::float(0.0)),
            positional_args
                .get(1)
                .cloned()
                .unwrap_or_else(|| Expression::float(0.0)),
        )),
        "real"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_array(&arg.value, type_env)) =>
        {
            let item_name = "__fortran_complex_item";
            Some(build_fortran_typed_array_map(
                args[0].value.clone(),
                fortran_complex_real_part(&Expression::ident(item_name)),
                false,
                item_name,
                "__fortran_complex_index",
                Some("complex".to_string()),
            ))
        }
        "real"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_scalar(&arg.value, type_env)) =>
        {
            Some(fortran_complex_real_part(&args[0].value))
        }
        "aimag"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_array(&arg.value, type_env)) =>
        {
            let item_name = "__fortran_complex_item";
            Some(build_fortran_typed_array_map(
                args[0].value.clone(),
                fortran_complex_imag_part(&Expression::ident(item_name)),
                false,
                item_name,
                "__fortran_complex_index",
                Some("complex".to_string()),
            ))
        }
        "aimag"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_scalar(&arg.value, type_env)) =>
        {
            Some(fortran_complex_imag_part(&args[0].value))
        }
        "conjg"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_array(&arg.value, type_env)) =>
        {
            let item_name = "__fortran_complex_item";
            Some(build_fortran_typed_array_map(
                args[0].value.clone(),
                build_fortran_complex_conjg_expr(&Expression::ident(item_name)),
                false,
                item_name,
                "__fortran_complex_index",
                Some("complex".to_string()),
            ))
        }
        "conjg"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_scalar(&arg.value, type_env)) =>
        {
            Some(build_fortran_complex_conjg_expr(&args[0].value))
        }
        "abs"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_array(&arg.value, type_env)) =>
        {
            let item_name = "__fortran_complex_item";
            Some(build_fortran_typed_array_map(
                args[0].value.clone(),
                build_fortran_complex_abs_expr(&Expression::ident(item_name)),
                false,
                item_name,
                "__fortran_complex_index",
                Some("complex".to_string()),
            ))
        }
        "abs"
            if args
                .first()
                .is_some_and(|arg| expr_is_fortran_complex_scalar(&arg.value, type_env)) =>
        {
            Some(build_fortran_complex_abs_expr(&args[0].value))
        }
        _ => None,
    }
}

fn lower_fortran_complex_binary_expr(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    type_env: &HashMap<String, String>,
) -> Expression {
    // A real operand mixed with a complex one is that real plus a zero
    // imaginary part — which is the only Fortran-specific thing here. The four
    // arithmetic rules themselves are the shared ones.
    let left_re = fortran_complex_real_or_scalar(left, type_env);
    let left_im = fortran_complex_imag_or_zero(left, type_env);
    let right_re = fortran_complex_real_or_scalar(right, type_env);
    let right_im = fortran_complex_imag_or_zero(right, type_env);
    match op {
        BinOp::Add => complex::add(left_re, left_im, right_re, right_im),
        BinOp::Sub => complex::sub(left_re, left_im, right_re, right_im),
        BinOp::Mul => complex::mul(left_re, left_im, right_re, right_im),
        BinOp::Div => complex::div(left_re, left_im, right_re, right_im),
        _ => Expression::null(),
    }
}

// ── Complex values are the SHARED representation ────────────────────────────
//
// `primitives::complex` is the model — an object `{real, imag}` that C, Python
// and Ruby already build and read. Fortran used to carry its own `{re, im}`
// with a private copy of the same arithmetic; these are the same operations,
// spelled once. The wrappers stay because the walker names them everywhere,
// and because `conjg`/`abs` read a VALUE while the shared builders take the two
// components apart.

/// The shared `{real, imag}` object, stamped `__type = "complex"`.
///
/// The stamp is what tells a DISPLAY path that this bag of two numbers is a
/// complex — `ObjectKind::Ordinary` alone cannot, because a derived-type value
/// is one too, which is why `[builtin_slots.object].to_string` is the wrong
/// instrument here. Python stamps the same shape for the same reason, and each
/// language renders it in its own spelling: `(1.0,2.0)` here, `(1+2j)` there.
fn build_fortran_complex_expr(real: Expression, imag: Expression) -> Expression {
    let ExprKind::Object(mut props) = complex::complex_object(real, imag).kind else {
        unreachable!("complex_object builds an object literal");
    };
    props.insert(
        0,
        ObjectProperty::KeyValue {
            key: Expression::string("__type"),
            value: Expression::string("complex"),
        },
    );
    Expression::new(ExprKind::Object(props))
}

fn build_fortran_complex_conjg_expr(value: &Expression) -> Expression {
    complex::conj(
        fortran_complex_real_part(value),
        fortran_complex_imag_part(value),
    )
}

fn build_fortran_complex_abs_expr(value: &Expression) -> Expression {
    complex::cabs(
        fortran_complex_real_part(value),
        fortran_complex_imag_part(value),
    )
}

fn fortran_complex_real_part(value: &Expression) -> Expression {
    complex::real_part(value.clone())
}

fn fortran_complex_imag_part(value: &Expression) -> Expression {
    complex::imag_part(value.clone())
}

fn fortran_complex_real_or_scalar(
    value: &Expression,
    type_env: &HashMap<String, String>,
) -> Expression {
    if expr_is_fortran_complex_scalar(value, type_env) {
        return fortran_complex_real_part(value);
    }
    value.clone()
}

fn fortran_complex_imag_or_zero(
    value: &Expression,
    type_env: &HashMap<String, String>,
) -> Expression {
    if expr_is_fortran_complex_scalar(value, type_env) {
        return fortran_complex_imag_part(value);
    }
    Expression::float(0.0)
}

fn lower_intrinsic_expr_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let lowered = name.to_ascii_lowercase();
    let positional_args = args
        .iter()
        .filter(|arg| {
            arg.name
                .as_deref()
                .is_none_or(|name| !name.eq_ignore_ascii_case("kind"))
        })
        .map(|arg| arg.value.clone())
        .collect::<Vec<_>>();
    match lowered.as_str() {
        "dot_product" if args.len() == 2 && args.iter().all(|arg| arg.name.is_none()) => Some(
            build_fortran_dot_product_expr(args[0].value.clone(), args[1].value.clone()),
        ),
        "transpose" if args.len() == 1 && args[0].name.is_none() => {
            Some(build_fortran_transpose_expr(args[0].value.clone()))
        }
        // Degree trig. The host math library is radians-only, so the conversion
        // IS the lowering — and the two families convert opposite sides:
        // `sind`/`cosd`/`tand` take degrees, so the ARGUMENT converts; the
        // inverse forms answer in degrees, so the RESULT does. Dropping the `d`
        // gives the radian name in both directions.
        "sind" | "cosd" | "tand" if positional_args.len() == 1 => {
            let radians = Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(positional_args[0].clone()),
                right: Box::new(Expression::float(std::f64::consts::PI / 180.0)),
            });
            Some(fortran_call(&lowered[..lowered.len() - 1], vec![radians]))
        }
        "asind" | "acosd" | "atand" if positional_args.len() == 1 => {
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(fortran_call(
                    &lowered[..lowered.len() - 1],
                    vec![positional_args[0].clone()],
                )),
                right: Box::new(Expression::float(180.0 / std::f64::consts::PI)),
            }))
        }
        // `new_line(x)` is the newline for `x`'s character kind — one character,
        // and the same one for every kind the AST represents. The argument only
        // selects the kind, so it is not evaluated.
        "new_line" if positional_args.len() == 1 => Some(Expression::string("\n")),
        // ── Single-image coarray inquiry ────────────────────────────────────
        //
        // Execution is one image, which the grammar has said all along ("walker
        // treats these as no-ops (single-image execution)") — it just never
        // answered the questions. Verified against
        // `gfortran -fcoarray=single`: image 1 of 1, and the INITIAL team is
        // numbered −1, not 1.
        // F2018 `reduce(array, operator(+) [, dim=] [, mask=] [, identity=]
        // [, ordered=])` is the general fold, and for the two operations the
        // suite uses it IS the intrinsic reduction that already exists — `+` is
        // `sum`, `*` is `product`, with the same `dim`/`mask` keywords. So it
        // rewrites rather than growing a second reduction path.
        //
        // `identity` only matters for an empty section (`sum`/`product` already
        // return 0/1 there) and `ordered` fixes an evaluation order that is
        // unobservable for associative `+`/`*` — both drop out.
        "reduce" if positional_args.len() >= 2 => {
            // The STANDARD form: OPERATION is a pure function's NAME.
            if let ExprKind::Ident(function) = &positional_args[1].kind {
                let named = |key: &str| {
                    args.iter()
                        .find(|a| a.name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(key)))
                        .map(|a| a.value.clone())
                };
                // `DIM` reduces along one axis and needs the ranked machinery;
                // leave it alone rather than answer it wrongly.
                if named("dim").is_some() {
                    return None;
                }
                return Some(build_fortran_reduce_call(
                    positional_args[0].clone(),
                    function,
                    named("mask"),
                    named("identity"),
                ));
            }
            let ExprKind::Lit(Literal::Str(op)) = &positional_args[1].kind else {
                return None;
            };
            let folded = match op.as_str() {
                "+" => "sum",
                "*" => "product",
                _ => return None,
            };
            let mut forwarded = vec![Argument::positional(positional_args[0].clone())];
            forwarded.extend(args.iter().filter(|arg| {
                arg.name
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case("dim") || name.eq_ignore_ascii_case("mask"))
            }).cloned());
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(folded)),
                args: forwarded,
                optional: false,
            }))
        }
        // `modulo(a, p)` takes the sign of the DIVISOR, `mod(a, p)` the sign of
        // the dividend — they differ exactly when the operands disagree in sign.
        // Both were mapped to the same truncating `f64_mod`, which is `mod`, so
        // `modulo(-10, 3)` answered −1 where gfortran answers 2.
        //
        // Built as the defining identity `a - floor(a/p) * p`. ⛔ The division
        // must be forced REAL: integer division truncates now, and truncation
        // is precisely what makes this `mod` again.
        "modulo" if positional_args.len() == 2 => {
            let real = |expr: Expression| {
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("real")),
                    args: vec![Argument::positional(expr)],
                    optional: false,
                })
            };
            let quotient = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("floor")),
                args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                    op: BinOp::Div,
                    left: Box::new(real(positional_args[0].clone())),
                    right: Box::new(real(positional_args[1].clone())),
                }))],
                optional: false,
            });
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(positional_args[0].clone()),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(quotient),
                    right: Box::new(positional_args[1].clone()),
                })),
            }))
        }
        "this_image" | "num_images" | "image_index" => Some(Expression::int(1)),
        "team_number" | "get_team" => Some(Expression::int(-1)),
        // Verified against `gfortran -fcoarray=single`: the one image is
        // healthy (`image_status` → 0 = STAT_OK) and neither list has members.
        "image_status" => Some(Expression::int(0)),
        "failed_images" | "stopped_images" => Some(fortran_array_expr(Vec::new())),
        // Every array vybe builds is one contiguous run — there is no strided
        // view to be non-contiguous about.
        "is_contiguous" => Some(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        // One image, one element per codimension.
        "coshape" => Some(fortran_array_expr(vec![Expression::int(1)])),
        // A coarray on one image has exactly one element per codimension, so
        // both cobounds are 1.
        "lcobound" | "ucobound" => Some(Expression::int(1)),
        // `radix` is the base of the numeric model: 2 for every integer and real
        // kind on every target vybe emits for. It takes no model lookup, unlike
        // `maxexponent`/`minexponent` which vary by kind.
        "radix" if positional_args.len() == 1 => Some(Expression::int(2)),
        // `INT(a)` TRUNCATES toward zero — `int(3.5)` is 3. A cast to "integer"
        // does not: it is a numeric coercion, and it answered 3.5. The profile
        // already declares `int = common:to_int`, which is the truncation, so
        // the only thing to do here is drop a `kind=` that has no bearing on
        // the value and let the builtin have the call. Folding it was what kept
        // the builtin from ever being reached — and what hid `int(whole_array)`
        // from the elementwise lowering, which handles `int` by name.
        "int" if positional_args.len() > 1 || args.len() > positional_args.len() => {
            Some(fortran_call("int", vec![positional_args.first()?.clone()]))
        }
        "real" | "dble" if !positional_args.is_empty() => {
            Some(Expression::new(ExprKind::Cast {
                expr: Box::new(positional_args[0].clone()),
                type_name: "number".to_string(),
            }))
        }
        "aint" if args.len() == 1 => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("trunc")),
            args: vec![Argument::positional(args[0].value.clone())],
            optional: false,
        })),
        // `nint(x)` is the INTEGER form of `anint(x)`, so it is spelled with
        // Fortran's own two intrinsics. It used to be intercepted by a
        // hardcoded `nint` arm in the SHARED compiler
        // (`primitives/builtins.rs`) that emitted ties-to-EVEN — a Fortran-only
        // name living in shared code, which also meant a user program defining
        // its own `nint` was hijacked. gfortran ties AWAY FROM ZERO.
        "nint" if args.len() == 1 => Some(fortran_call(
            "int",
            vec![fortran_call("round", vec![args[0].value.clone()])],
        )),
        "anint" if args.len() == 1 => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("round")),
            args: vec![Argument::positional(args[0].value.clone())],
            optional: false,
        })),
        "achar" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("char")),
            args: args.to_vec(),
            optional: false,
        })),
        "iachar" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("ichar")),
            args: args.to_vec(),
            optional: false,
        })),
        "associated" if args.len() == 2 => {
            let not_null = Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(args[0].value.clone()),
                right: Box::new(Expression::null()),
            });
            let same_target = Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(args[0].value.clone()),
                right: Box::new(args[1].value.clone()),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(not_null),
                    right: Box::new(same_target),
                })),
                then: Box::new(Expression::bool(true)),
                else_: Box::new(Expression::bool(false)),
            }))
        }
        "associated" | "allocated" if args.len() == 1 => {
            // Produce a JS boolean (true/false) not an i32 (1/0) so print * formats correctly.
            let not_null = Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(args[0].value.clone()),
                right: Box::new(Expression::null()),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(not_null),
                then: Box::new(Expression::bool(true)),
                else_: Box::new(Expression::bool(false)),
            }))
        }
        "len" if args.len() == 1 => Some(Expression::new(ExprKind::Member {
            object: Box::new(args[0].value.clone()),
            field: "length".to_string(),
            null_safe: false,
        })),
        // `rank` is NOT rewritten here: the profile binds the spelling straight
        // to `common:collections.rank`, the way `matmul` and `size` are bound.
        // Rewriting it to a call to itself would either loop this pass or
        // re-enter it for no gain.
        "transfer" if positional_args.len() >= 2 => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__fortran_transfer")),
            args: args.to_vec(),
            optional: false,
        })),
        "dim" if args.len() == 2 => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("max")),
            args: vec![
                Argument::positional(Expression::new(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(args[0].value.clone()),
                    right: Box::new(args[1].value.clone()),
                })),
                Argument::positional(Expression::int(0)),
            ],
            optional: false,
        })),
        // `MERGE` is elementwise SELECT — the shared ranked-array node beside
        // PACK/UNPACK/RESHAPE, not a Fortran map. The private builder it
        // replaced mapped over the mask at rank 1, so a rank-2 mask handed the
        // callback a ROW; a row is always truthy, and every element took the
        // true branch.
        "merge" if args.len() == 3 => {
            if let Some(value) = fortran_expr_is_literal_bool(&args[2].value) {
                // A literal scalar mask has no shape to walk and folds outright.
                Some(Expression::new(ExprKind::Ternary {
                    cond: Box::new(Expression::bool(value)),
                    then: Box::new(args[0].value.clone()),
                    else_: Box::new(args[1].value.clone()),
                }))
            } else {
                Some(fortran_merge_node(
                    args[0].value.clone(),
                    args[1].value.clone(),
                    args[2].value.clone(),
                ))
            }
        }
        "llt" if args.len() == 2 => Some(build_fortran_lexical_compare_expr(
            BinOp::Lt,
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "lle" if args.len() == 2 => Some(build_fortran_lexical_compare_expr(
            BinOp::LtEq,
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "lgt" if args.len() == 2 => Some(build_fortran_lexical_compare_expr(
            BinOp::Gt,
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "lge" if args.len() == 2 => Some(build_fortran_lexical_compare_expr(
            BinOp::GtEq,
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "cshift" if positional_args.len() >= 2 => Some(build_fortran_cshift_1d_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
        )),
        "eoshift" if positional_args.len() >= 2 => {
            let boundary = positional_args
                .get(2)
                .cloned()
                .or_else(|| {
                    args.iter()
                        .find(|arg| {
                            arg.name
                                .as_deref()
                                .is_some_and(|name| name.eq_ignore_ascii_case("boundary"))
                        })
                        .map(|arg| arg.value.clone())
                })
                .unwrap_or_else(|| Expression::int(0));
            Some(build_fortran_eoshift_1d_expr(
                positional_args[0].clone(),
                positional_args[1].clone(),
                boundary,
            ))
        }
        "pack" if positional_args.len() >= 2 => {
            let mut transform_args = vec![positional_args[0].clone(), positional_args[1].clone()];
            if let Some(vector) = positional_args.get(2).cloned().or_else(|| {
                args.iter()
                    .find(|arg| {
                        arg.name
                            .as_deref()
                            .is_some_and(|name| name.eq_ignore_ascii_case("vector"))
                    })
                    .map(|arg| arg.value.clone())
            }) {
                transform_args.push(vector);
            }
            Some(Expression::new(ExprKind::ArrayTransform {
                op: ArrayTransformOp::PackMask,
                args: transform_args,
                order: ArrayTraversalOrder::ColumnMajor,
            }))
        }
        "unpack" if positional_args.len() >= 3 => Some(Expression::new(ExprKind::ArrayTransform {
            op: ArrayTransformOp::UnpackMask,
            args: vec![
                positional_args[0].clone(),
                positional_args[1].clone(),
                positional_args[2].clone(),
            ],
            order: ArrayTraversalOrder::ColumnMajor,
        })),
        "spread" if positional_args.len() == 3 => {
            let dim = fortran_literal_int(&positional_args[1])?;
            let source = positional_args[0].clone();
            let ncopies = positional_args[2].clone();
            match dim {
                1 => Some(build_fortran_spread_dim1_expr(source, ncopies)),
                2 => Some(build_fortran_spread_dim2_expr(source, ncopies)),
                _ => None,
            }
        }
        // `RESHAPE(SOURCE, SHAPE [, PAD] [, ORDER])`. All four are ordinary
        // arguments, so `pad` and `order` are read by keyword OR by position —
        // and the third position is PAD, which is what the earlier lowering had
        // wrong. `order` is a permutation of the dimensions, never the string
        // `"C"`; the identity is Fortran's own element order and the full
        // reversal is C's, which are the two the shared node names.
        "reshape" if positional_args.len() >= 2 => {
            let mut transform_args =
                vec![positional_args[0].clone(), positional_args[1].clone()];
            if let Some(pad) = fortran_argument("pad", 2, args, &positional_args) {
                transform_args.push(pad);
            }
            let order = match fortran_argument("order", 3, args, &positional_args) {
                Some(order) => fortran_traversal_order(&order)?,
                None => ArrayTraversalOrder::ColumnMajor,
            };
            Some(Expression::new(ExprKind::ArrayTransform {
                op: ArrayTransformOp::Reshape,
                args: transform_args,
                order,
            }))
        }
        "sign" if args.len() == 2 => Some(build_fortran_sign_expr(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "hypot" if args.len() == 2 => Some(build_fortran_hypot_expr(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        // `bessel_jn(n1, n2, x)` / `bessel_yn(n1, n2, x)` — the TRANSFORMATIONAL
        // form, which returns the whole run of orders n1..n2 as an array rather
        // than one value. It is the same shape as an implied-do array
        // constructor, so it lowers to exactly that: `[(bessel_jn(i, x),
        // i = n1, n2)]`. No new machinery, and the scalar form stays untouched.
        "bessel_jn" | "bessel_yn" if args.len() == 3 => {
            Some(build_fortran_bessel_series_expr(
                &lowered,
                args[0].value.clone(),
                args[1].value.clone(),
                args[2].value.clone(),
            ))
        }
        // F2008 `norm2(a)` is the Euclidean norm — `sqrt(sum(a*a))`. Both
        // halves already exist, so this is a spelling, not a new primitive.
        // `norm2(a, dim)` folds along one dimension and passes `dim` straight
        // through to the same `sum`.
        "norm2" if !args.is_empty() && args.len() <= 2 => {
            Some(build_fortran_norm2_expr(&args))
        }
        "mod" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(args[0].value.clone()),
            right: Box::new(args[1].value.clone()),
        })),
        "modulo" if args.len() == 2 => Some(build_fortran_modulo_expr(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "iand" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(args[0].value.clone()),
            right: Box::new(args[1].value.clone()),
        })),
        "ior" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitOr,
            left: Box::new(args[0].value.clone()),
            right: Box::new(args[1].value.clone()),
        })),
        "ieor" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitXor,
            left: Box::new(args[0].value.clone()),
            right: Box::new(args[1].value.clone()),
        })),
        "not" if args.len() == 1 => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr: Box::new(args[0].value.clone()),
        })),
        "ishft" if args.len() == 2 => Some(build_fortran_ishft_expr(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        // F2008's directional shifts. Unlike `ishft` the direction is in the
        // NAME, so the shift count is always non-negative and no branch is
        // needed. `shiftr` is LOGICAL (zero fill) and `shifta` ARITHMETIC (sign
        // fill) — the one place the two differ, and the reason they are separate
        // intrinsics at all.
        "shiftl" | "shiftr" | "shifta" if positional_args.len() == 2 => {
            Some(Expression::new(ExprKind::Binary {
                op: match lowered.as_str() {
                    "shiftl" => BinOp::Shl,
                    "shiftr" => BinOp::UShr,
                    _ => BinOp::Shr,
                },
                left: Box::new(positional_args[0].clone()),
                right: Box::new(positional_args[1].clone()),
            }))
        }
        // `poppar` is the PARITY of the population count — 1 when an odd number
        // of bits is set. Built on `popcnt`, which the profile already emits, so
        // the bit counting is not written twice.
        "poppar" if positional_args.len() == 1 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(fortran_call("popcnt", vec![positional_args[0].clone()])),
            right: Box::new(Expression::int(1)),
        })),
        "ibset" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitOr,
            left: Box::new(args[0].value.clone()),
            right: Box::new(build_fortran_bit_mask(args[1].value.clone())),
        })),
        "ibclr" if args.len() == 2 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(args[0].value.clone()),
            right: Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(build_fortran_bit_mask(args[1].value.clone())),
            })),
        })),
        "btest" if positional_args.len() == 2 => Some(build_fortran_btest_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
        )),
        // `ishftc(i, shift)` with no size rotates the whole integer.
        "ishftc" if positional_args.len() >= 2 => Some(build_fortran_ishftc_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
            positional_args
                .get(2)
                .cloned()
                .unwrap_or_else(|| Expression::int(FORTRAN_BIT_SIZE)),
        )),
        "maskl" if !positional_args.is_empty() => {
            Some(build_fortran_maskl_expr(positional_args[0].clone()))
        }
        "maskr" if !positional_args.is_empty() => {
            Some(build_fortran_maskr_expr(positional_args[0].clone()))
        }
        "merge_bits" if positional_args.len() == 3 => Some(build_fortran_merge_bits_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
            positional_args[2].clone(),
        )),
        "dshiftl" if positional_args.len() == 3 => Some(build_fortran_dshiftl_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
            positional_args[2].clone(),
        )),
        "dshiftr" if positional_args.len() == 3 => Some(build_fortran_dshiftr_expr(
            positional_args[0].clone(),
            positional_args[1].clone(),
            positional_args[2].clone(),
        )),
        "bge" | "bgt" | "ble" | "blt" if positional_args.len() == 2 => {
            let op = match lowered.as_str() {
                "bge" => BinOp::GtEq,
                "bgt" => BinOp::Gt,
                "ble" => BinOp::LtEq,
                _ => BinOp::Lt,
            };
            Some(build_fortran_bit_compare_expr(
                op,
                positional_args[0].clone(),
                positional_args[1].clone(),
            ))
        }
        // `parity` is deliberately NOT folded here — see the profile. A fold
        // fires on the NAME, before anything knows whether the program defines
        // a function of its own by that name, and the suite has a contained
        // `recursive integer function parity(n, even_only)` that a fold would
        // silently hijack. Profile builtins lose to a user definition, which is
        // what Fortran requires of an intrinsic.
        // Both used to answer a flat `8`, so `selected_int_kind(123)` claimed a
        // kind that holds 123 decimal digits existed.
        "selected_int_kind" if !positional_args.is_empty() => {
            fortran_const_int(&positional_args[0])
                .map(|range| Expression::int(fortran_selected_int_kind(range)))
        }
        "selected_real_kind" if !positional_args.is_empty() => {
            let precision = fortran_const_int(&positional_args[0])?;
            // The range argument is optional: `selected_real_kind(6)` asks for
            // precision alone.
            let range = match positional_args.get(1) {
                Some(arg) => fortran_const_int(arg)?,
                None => 0,
            };
            Some(Expression::int(fortran_selected_real_kind(precision, range)))
        }
        "kind" if args.len() == 1 => {
            fold_fortran_type_inquiry("kind", &positional_args[0], &HashMap::new(), None)
        }
        "bit_size" if args.len() == 1 => {
            fold_fortran_type_inquiry("bit_size", &positional_args[0], &HashMap::new(), None)
        }
        "storage_size" if !positional_args.is_empty() => {
            fold_fortran_type_inquiry("storage_size", &positional_args[0], &HashMap::new(), None)
        }
        "precision" if args.len() == 1 => {
            fold_fortran_type_inquiry("precision", &positional_args[0], &HashMap::new(), None)
        }
        "range" if args.len() == 1 => {
            fold_fortran_type_inquiry("range", &positional_args[0], &HashMap::new(), None)
        }
        "digits" if args.len() == 1 => {
            fold_fortran_type_inquiry("digits", &positional_args[0], &HashMap::new(), None)
        }
        "maxexponent" | "minexponent" if args.len() == 1 => {
            fold_fortran_type_inquiry(&lowered, &positional_args[0], &HashMap::new(), None)
        }
        "huge" if args.len() == 1 => Some(build_fortran_huge_expr(&args[0].value)),
        "tiny" if args.len() == 1 => Some(Expression::float(f32::MIN_POSITIVE as f64)),
        "epsilon" if args.len() == 1 => Some(Expression::float(f32::EPSILON as f64)),

        // ── IEEE_ARITHMETIC and the real-number model ──────────────────────
        //
        // `ieee_is_nan` and `ieee_is_finite` are profile builtins — ECMA asks
        // the same two questions. What is left either needs a value no host
        // function returns (an infinity, a class code) or decomposes into
        // arithmetic, and both are expressions the AST can already say.
        "ieee_value" if args.len() == 2 => fortran_ieee_class_constant(&args[1].value),
        "ieee_class" if args.len() == 1 => Some(build_fortran_ieee_class_expr(&args[0].value)),
        "ieee_is_normal" if args.len() == 1 => {
            Some(build_fortran_ieee_is_normal_expr(&args[0].value))
        }
        "ieee_unordered" if args.len() == 2 => Some(fortran_bin(
            BinOp::Or,
            fortran_ieee_is_nan(args[0].value.clone()),
            fortran_ieee_is_nan(args[1].value.clone()),
        )),
        "ieee_copy_sign" if args.len() == 2 => Some(build_fortran_ieee_copy_sign_expr(
            &args[0].value,
            &args[1].value,
        )),
        "ieee_logb" if args.len() == 1 => Some(fortran_bin(
            BinOp::Sub,
            build_fortran_exponent_expr(&args[0].value),
            Expression::int(1),
        )),
        "ieee_rint" if args.len() == 1 => Some(fortran_call("round", vec![args[0].value.clone()])),
        "ieee_scalb" | "scale" if args.len() == 2 => Some(build_fortran_scale_expr(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "ieee_rem" if args.len() == 2 => {
            let quotient = fortran_bin(
                BinOp::Div,
                args[0].value.clone(),
                args[1].value.clone(),
            );
            Some(fortran_bin(
                BinOp::Sub,
                args[0].value.clone(),
                fortran_bin(
                    BinOp::Mul,
                    args[1].value.clone(),
                    fortran_call("round", vec![quotient]),
                ),
            ))
        }
        // Every kind this compiler has is supported, and none of them signal.
        name if name.starts_with("ieee_support_") => {
            Some(Expression::new(ExprKind::Lit(Literal::Bool(true))))
        }
        "exponent" if args.len() == 1 => Some(build_fortran_exponent_expr(&args[0].value)),
        "fraction" if args.len() == 1 => Some(build_fortran_scale_expr(
            args[0].value.clone(),
            fortran_unary_minus(build_fortran_exponent_expr(&args[0].value)),
        )),
        // `nearest`/`spacing` are folded in `lower_fortran_type_inquiry_in_expr`
        // instead: the LANE depends on the operand's declared kind, and that is
        // the only pass carrying a `type_env`.
        "rrspacing" if args.len() == 1 => Some(fortran_bin(
            BinOp::Div,
            fortran_call("abs", vec![args[0].value.clone()]),
            fortran_call("spacing", vec![args[0].value.clone()]),
        )),
        "set_exponent" if args.len() == 2 => Some(build_fortran_scale_expr(
            build_fortran_scale_expr(
                args[0].value.clone(),
                fortran_unary_minus(build_fortran_exponent_expr(&args[0].value)),
            ),
            args[1].value.clone(),
        )),
        // ── Reductions with `dim=` — ONE lowering for the whole family ──────
        //
        // Each is the scalar reduction it already has, mapped over the lanes of
        // the named dimension. They sit ahead of the whole-array arms below
        // because those match on `args.len() == 1` and a `dim` is a second
        // argument, which is what used to leave every one of these calling a
        // function that does not exist.
        "all" | "any" | "count" | "sum" | "product" | "maxval" | "minval"
            if fortran_dim_argument(args).is_some() =>
        {
            let dim = fortran_dim_argument(args)?;
            let array = args[0].value.clone();
            let name = lowered.clone();
            build_fortran_dim_reduction_expr(array, dim, move |lane| match name.as_str() {
                "all" => build_fortran_logical_array_reducer(lane, "every"),
                "any" => build_fortran_logical_array_reducer(lane, "some"),
                "count" => build_fortran_count_expr(lane),
                "product" => build_fortran_product_expr(lane),
                "minval" => build_fortran_nested_array_reduction("min", lane, 1, 0),
                "maxval" => build_fortran_nested_array_reduction("max", lane, 1, 0),
                _ => build_fortran_nested_array_reduction("sum", lane, 1, 0),
            })
        }
        "all" if args.len() == 1 => Some(build_fortran_logical_array_reducer(
            args[0].value.clone(),
            "every",
        )),
        "any" if args.len() == 1 => Some(build_fortran_logical_array_reducer(
            args[0].value.clone(),
            "some",
        )),
        "count" if args.len() == 1 => Some(build_fortran_count_expr(args[0].value.clone())),
        "product" if args.len() == 1 && args.iter().all(|arg| arg.name.is_none()) => {
            Some(build_fortran_product_expr(args[0].value.clone()))
        }
        "verify" if args.len() >= 2 => {
            let mut lowered_args = args.to_vec();
            lowered_args[0] = Argument::positional(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("trim")),
                args: vec![Argument::positional(args[0].value.clone())],
                optional: false,
            }));
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("verify")),
                args: lowered_args,
                optional: false,
            }))
        }
        _ => None,
    }
}

/// The `dim` of a reduction call, when it is a literal.
///
/// `sum(a, dim, mask)` is the shape of every one of them — `all(m, dim)`,
/// `count(m, dim)`, `maxval(m, dim)` — so `dim` is the second POSITIONAL or a
/// `dim=` keyword. A non-literal dim gets no fold: which dimension is being
/// reduced decides the shape of the traversal, and that is a compile-time
/// question here.
fn fortran_dim_argument(args: &[Argument]) -> Option<i64> {
    let named = args.iter().find(|arg| {
        arg.name
            .as_deref()
            .is_some_and(|name| name.eq_ignore_ascii_case("dim"))
    });
    let value = match named {
        Some(arg) => &arg.value,
        None => {
            let second = args.get(1)?;
            if second.name.is_some() {
                return None;
            }
            &second.value
        }
    };
    fortran_literal_int(value)
}

/// `<reduction>(array, dim=N)` — reduce ONE dimension of a rank-2 array,
/// yielding a rank-1 array.
///
/// A rank-2 array is stored as an array of ROWS — `build_fortran_reshape_2d_expr`
/// maps dim1 outside and dim2 inside, so `m(i,j)` is `m[i][j]`. Reducing along
/// dim 2 therefore maps the scalar reduction over the rows exactly as they sit,
/// and dim 1 — which varies the FIRST subscript, i.e. walks columns — is the
/// same map over the TRANSPOSE. Both halves already exist; this is the map
/// between them, not a new traversal.
fn build_fortran_dim_reduction_expr(
    array: Expression,
    dim: i64,
    scalar_reduce: impl Fn(Expression) -> Expression,
) -> Option<Expression> {
    let lanes = match dim {
        1 => build_fortran_transpose_expr(array),
        2 => array,
        // Rank 3 and up needs a lane-walk this shape cannot express.
        _ => return None,
    };
    let lane_name = "__fortran_dim_reduce_lane";
    Some(build_fortran_array_map(
        lanes,
        scalar_reduce(Expression::ident(lane_name)),
        false,
        lane_name,
        "__fortran_dim_reduce_index",
    ))
}

/// `product(array)` — the scalar form, shared by the whole-array arm and by
/// each lane of the `dim=` form.
fn build_fortran_product_expr(array: Expression) -> Expression {
    let acc_name = "__fortran_product_acc";
    let item_name = "__fortran_product_item";
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array),
            field: "reduce".to_string(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Lambda {
                params: vec![
                    Param {
                        name: acc_name.to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    },
                    Param {
                        name: item_name.to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    },
                ],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(Expression::ident(acc_name)),
                    right: Box::new(Expression::ident(item_name)),
                }))),
                is_async: false,
                captures: Vec::new(),
            })),
            Argument::positional(Expression::int(1)),
        ],
        optional: false,
    })
}

fn build_fortran_bit_mask(shift: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(Expression::int(1)),
        right: Box::new(shift),
    })
}

/// Default INTEGER is kind 4 — 32 bits. `bit_size(0)` already answers 32
/// (`fold_fortran_type_inquiry`), and the bit model here has to agree with it.
const FORTRAN_BIT_SIZE: i64 = 32;

fn fortran_bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn fortran_bit_not(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::BitNot,
        expr: Box::new(expr),
    })
}


fn fortran_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn fortran_unary_minus(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Neg,
        expr: Box::new(expr),
    })
}

fn fortran_ieee_is_nan(value: Expression) -> Expression {
    fortran_call("ieee_is_nan", vec![value])
}

/// The `IEEE_ARITHMETIC` class constants. `ieee_value(x, IEEE_POSITIVE_INF)` is
/// the only way to say "infinity" in Fortran — there is no literal for it — so
/// the constant is read as a spelling here and answered with the value itself.
fn fortran_ieee_class_constant(class: &Expression) -> Option<Expression> {
    let ExprKind::Ident(name) = &class.kind else {
        return None;
    };
    let value = match name.to_ascii_lowercase().as_str() {
        "ieee_positive_inf" => f64::INFINITY,
        "ieee_negative_inf" => f64::NEG_INFINITY,
        "ieee_quiet_nan" | "ieee_signaling_nan" => f64::NAN,
        "ieee_positive_zero" => 0.0,
        "ieee_negative_zero" => -0.0,
        "ieee_positive_normal" => 1.0,
        "ieee_negative_normal" => -1.0,
        "ieee_positive_denormal" | "ieee_positive_subnormal" => f32::MIN_POSITIVE as f64 / 2.0,
        "ieee_negative_denormal" | "ieee_negative_subnormal" => -(f32::MIN_POSITIVE as f64) / 2.0,
        _ => return None,
    };
    Some(Expression::float(value))
}

/// `ieee_class(x)` answers with one of the same constants, so it is built from
/// the values they carry — a comparison chain, not a table of codes.
fn build_fortran_ieee_class_expr(value: &Expression) -> Expression {
    let is_negative = fortran_bin(BinOp::Lt, value.clone(), Expression::float(0.0));
    let signed = |positive: f64, negative: f64| {
        fortran_ternary(
            is_negative.clone(),
            Expression::float(negative),
            Expression::float(positive),
        )
    };
    fortran_ternary(
        fortran_ieee_is_nan(value.clone()),
        Expression::float(f64::NAN),
        fortran_ternary(
            fortran_call("ieee_is_finite", vec![value.clone()]),
            fortran_ternary(
                fortran_bin(BinOp::Eq, value.clone(), Expression::float(0.0)),
                signed(0.0, -0.0),
                fortran_ternary(
                    fortran_bin(
                        BinOp::Lt,
                        fortran_call("abs", vec![value.clone()]),
                        Expression::float(f32::MIN_POSITIVE as f64),
                    ),
                    signed(
                        f32::MIN_POSITIVE as f64 / 2.0,
                        -(f32::MIN_POSITIVE as f64) / 2.0,
                    ),
                    signed(1.0, -1.0),
                ),
            ),
            signed(f64::INFINITY, f64::NEG_INFINITY),
        ),
    )
}

/// Normal means finite, non-zero, and at least `tiny` in magnitude — the three
/// things a subnormal, a zero, an infinity and a NaN each fail.
fn build_fortran_ieee_is_normal_expr(value: &Expression) -> Expression {
    fortran_bin(
        BinOp::And,
        fortran_call("ieee_is_finite", vec![value.clone()]),
        fortran_bin(
            BinOp::GtEq,
            fortran_call("abs", vec![value.clone()]),
            Expression::float(f32::MIN_POSITIVE as f64),
        ),
    )
}

fn build_fortran_ieee_copy_sign_expr(magnitude: &Expression, sign: &Expression) -> Expression {
    fortran_ternary(
        fortran_bin(BinOp::Lt, sign.clone(), Expression::float(0.0)),
        fortran_unary_minus(fortran_call("abs", vec![magnitude.clone()])),
        fortran_call("abs", vec![magnitude.clone()]),
    )
}

/// `exponent(x)` — the power of two `x` sits above, one-based as Fortran counts
/// it, and zero for a zero operand.
fn build_fortran_exponent_expr(value: &Expression) -> Expression {
    fortran_ternary(
        fortran_bin(BinOp::Eq, value.clone(), Expression::float(0.0)),
        Expression::int(0),
        fortran_bin(
            BinOp::Add,
            fortran_call(
                "floor",
                vec![fortran_call(
                    "__fortran_log2",
                    vec![fortran_call("abs", vec![value.clone()])],
                )],
            ),
            Expression::int(1),
        ),
    )
}

/// `scale(x, i)` — `x * 2**i`, which is also `ieee_scalb` and the engine behind
/// `fraction` and `set_exponent`.
fn build_fortran_scale_expr(value: Expression, by: Expression) -> Expression {
    fortran_bin(
        BinOp::Mul,
        value,
        fortran_bin(BinOp::Pow, Expression::float(2.0), by),
    )
}

/// `spacing(x)` — the distance to the next representable number of the same
/// kind, `2**(exponent(x) - digits)`.
// `build_fortran_spacing_expr` computed the ULP from the EXPONENT. It is
// `common:math.ulp` now, derived from the adjacent representable value, which
// is right at a binade boundary by construction.

fn fortran_ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_),
    })
}

/// `btest(i, pos)` — is bit `pos` of `i` set?
///
/// The shift is LOGICAL: `btest(-1, 31)` asks about the sign bit, and an
/// arithmetic shift would smear the sign across it and answer for bit 31 of a
/// value that no longer has one.
fn build_fortran_btest_expr(value: Expression, pos: Expression) -> Expression {
    fortran_bin(
        BinOp::NotEq,
        fortran_bin(
            BinOp::BitAnd,
            fortran_bin(BinOp::UShr, value, pos),
            Expression::int(1),
        ),
        Expression::int(0),
    )
}

/// `maskr(n)` — the `n` rightmost bits set.
///
/// `1 << 32` is `1`, not `0`: WASM and ECMA both take the shift count mod 32.
/// The full-width case therefore cannot go through the shift at all — hence the
/// ternary rather than a plain `(1 << n) - 1`. `maskr(0)` needs no such guard,
/// `1 << 0 - 1` being 0 already.
fn build_fortran_maskr_expr(bits: Expression) -> Expression {
    fortran_ternary(
        fortran_bin(
            BinOp::GtEq,
            bits.clone(),
            Expression::int(FORTRAN_BIT_SIZE),
        ),
        Expression::int(-1),
        fortran_bin(
            BinOp::Sub,
            fortran_bin(BinOp::Shl, Expression::int(1), bits),
            Expression::int(1),
        ),
    )
}

/// `maskl(n)` — the `n` leftmost bits set.
///
/// The mirror of `maskr`'s problem lands on zero instead of full width:
/// `-1 << (32 - 0)` is `-1`, where the answer is 0.
fn build_fortran_maskl_expr(bits: Expression) -> Expression {
    fortran_ternary(
        fortran_bin(BinOp::LtEq, bits.clone(), Expression::int(0)),
        Expression::int(0),
        fortran_bin(
            BinOp::Shl,
            Expression::int(-1),
            fortran_bin(BinOp::Sub, Expression::int(FORTRAN_BIT_SIZE), bits),
        ),
    )
}

/// `ishftc(i, shift [, size])` — rotate the rightmost `size` bits, leaving the
/// rest of `i` alone. `shift` may be negative, which rotates right.
///
/// The rotation amount is `modulo`, NOT `mod`: Fortran's `mod` keeps the
/// dividend's sign, so `mod(-2, 4)` is `-2` and a right rotate would come out
/// as a negative shift count. `modulo(-2, 4)` is 2 — the equivalent left
/// rotate, which is what a rotation by a negative amount means.
fn build_fortran_ishftc_expr(
    value: Expression,
    shift: Expression,
    size: Expression,
) -> Expression {
    // A full-width rotation IS `i32.rotl` — one instruction, and immune to this
    // region's `ShiftOverflow::Zero`, which the mask-and-shift lowering below is
    // not: it leans on `field >>> 32` being `field`, and under Fortran's own
    // shift rule that expression is 0.
    if fortran_literal_int(&size) == Some(FORTRAN_BIT_SIZE) {
        return fortran_bin(
            BinOp::RotL(vybe_ast::BitLane::W32),
            value,
            build_fortran_modulo_expr(shift, size),
        );
    }
    let mask = build_fortran_maskr_expr(size.clone());
    let amount = build_fortran_modulo_expr(shift, size.clone());
    let field = fortran_bin(BinOp::BitAnd, value.clone(), mask.clone());
    let rotated = fortran_bin(
        BinOp::BitAnd,
        fortran_bin(
            BinOp::BitOr,
            fortran_bin(BinOp::Shl, field.clone(), amount.clone()),
            fortran_bin(BinOp::UShr, field, fortran_bin(BinOp::Sub, size, amount)),
        ),
        mask.clone(),
    );
    // A zero rotation of a full-width field is the one case where the two
    // halves overlap — `field >>> 32` is `field` — and it is also the case
    // where OR-ing them is harmless, so no extra guard is needed.
    fortran_bin(
        BinOp::BitOr,
        fortran_bin(BinOp::BitAnd, value, fortran_bit_not(mask)),
        rotated,
    )
}

/// `dshiftl(i, j, shift)` — the leftmost `shift` bits of `j` become the
/// rightmost bits of the result; the rest is `i` shifted left to make room.
///
/// Both endpoints have to sidestep the shift: at 0 the `j` half would be
/// `j >>> 32` and at full width the `i` half would be `i << 32`, and a shift
/// count of 32 is a shift of 0.
fn build_fortran_dshiftl_expr(i: Expression, j: Expression, shift: Expression) -> Expression {
    let combined = fortran_bin(
        BinOp::BitOr,
        fortran_bin(BinOp::Shl, i.clone(), shift.clone()),
        fortran_bin(
            BinOp::UShr,
            j.clone(),
            fortran_bin(
                BinOp::Sub,
                Expression::int(FORTRAN_BIT_SIZE),
                shift.clone(),
            ),
        ),
    );
    fortran_ternary(
        fortran_bin(BinOp::Eq, shift.clone(), Expression::int(0)),
        i,
        fortran_ternary(
            fortran_bin(BinOp::GtEq, shift, Expression::int(FORTRAN_BIT_SIZE)),
            j,
            combined,
        ),
    )
}

/// `dshiftr(i, j, shift)` — the mirror: the rightmost `shift` bits of `i`
/// become the leftmost bits of the result.
fn build_fortran_dshiftr_expr(i: Expression, j: Expression, shift: Expression) -> Expression {
    let combined = fortran_bin(
        BinOp::BitOr,
        fortran_bin(
            BinOp::Shl,
            i.clone(),
            fortran_bin(
                BinOp::Sub,
                Expression::int(FORTRAN_BIT_SIZE),
                shift.clone(),
            ),
        ),
        fortran_bin(BinOp::UShr, j.clone(), shift.clone()),
    );
    fortran_ternary(
        fortran_bin(BinOp::Eq, shift.clone(), Expression::int(0)),
        j,
        fortran_ternary(
            fortran_bin(BinOp::GtEq, shift, Expression::int(FORTRAN_BIT_SIZE)),
            i,
            combined,
        ),
    )
}

/// `merge_bits(i, j, mask)` — bits of `i` where the mask is set, bits of `j`
/// where it is clear.
fn build_fortran_merge_bits_expr(
    i: Expression,
    j: Expression,
    mask: Expression,
) -> Expression {
    fortran_bin(
        BinOp::BitOr,
        fortran_bin(BinOp::BitAnd, i, mask.clone()),
        fortran_bin(BinOp::BitAnd, j, fortran_bit_not(mask)),
    )
}

/// `bge`/`bgt`/`ble`/`blt` — compare two integers as BIT SEQUENCES, i.e.
/// unsigned. `bgt(-1, 1)` is true because `-1` is all ones.
///
/// Flipping the sign bit of both operands maps unsigned order onto signed
/// order: everything with the top bit set sorts above everything without it,
/// and the remaining 31 bits already compare correctly. The obvious spelling —
/// ToUint32 both sides with `>>> 0` and compare — produces 4294967295, which
/// does not survive Fortran's integer comparison (it comes back as −1 and
/// answers false), so the comparison has to stay inside the signed range.
fn build_fortran_bit_compare_expr(op: BinOp, i: Expression, j: Expression) -> Expression {
    let sign_bit = || Expression::int(i32::MIN as i64);
    fortran_bin(
        op,
        fortran_bin(BinOp::BitXor, i, sign_bit()),
        fortran_bin(BinOp::BitXor, j, sign_bit()),
    )
}


/// `norm2(a[, dim])` → `sqrt(sum(a * a[, dim]))`.
fn build_fortran_norm2_expr(args: &[Argument]) -> Expression {
    let array = args[0].value.clone();
    let squared = Expression::new(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(array.clone()),
        right: Box::new(array),
    });
    let mut sum_args = vec![Argument::positional(squared)];
    if let Some(dim) = args.get(1) {
        sum_args.push(dim.clone());
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("sqrt")),
        args: vec![Argument::positional(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("sum")),
            args: sum_args,
            optional: false,
        }))],
        optional: false,
    })
}

/// `bessel_jn(n1, n2, x)` → the array `[(bessel_jn(i, x), i = n1, n2)]`.
///
/// Built through the same trip-count/map pair an implied-do constructor uses,
/// so the element expression is evaluated once per order with the loop index
/// substituted in.
fn build_fortran_bessel_series_expr(
    name: &str,
    first: Expression,
    last: Expression,
    x: Expression,
) -> Expression {
    let index_name = "__fortran_array_index";
    let step = Expression::int(1);
    let size = build_fortran_implied_do_trip_count(first.clone(), last, step.clone());
    let order = build_fortran_implied_do_value(first, step, index_name);
    let element = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: vec![Argument::positional(order), Argument::positional(x)],
        optional: false,
    });
    let array_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(size),
            Argument::positional(Expression::int(0)),
        ],
        optional: false,
    });
    build_fortran_array_map(
        array_expr,
        element,
        true,
        "__fortran_array_item",
        index_name,
    )
}

fn build_fortran_hypot_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("sqrt")),
        args: vec![Argument::positional(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(left.clone()),
                right: Box::new(left),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(right.clone()),
                right: Box::new(right),
            })),
        }))],
        optional: false,
    })
}

fn build_fortran_transfer_expr_with_hint(
    source: Expression,
    mold: Expression,
    size: Option<Expression>,
    target_hint: Option<&str>,
    source_hint: Option<&str>,
) -> Expression {
    if fortran_type_hint_is_array(target_hint) {
        return build_fortran_transfer_array_expr(source, size, source_hint, target_hint);
    }
    if source_hint.is_some_and(is_fortran_complex_type_hint) && !fortran_type_hint_is_array(target_hint)
    {
        return build_fortran_complex_expr_from_array(source);
    }
    // `transfer(x, 0)` / `transfer(n, 0.0)` reinterpret the BITS: the answer to
    // `transfer(3.14, 0)` is 1078523331, not 3. Bound to the same shared
    // primitives Go reaches through `math.Float32bits`.
    // TRANSFER is a bit cast, and a bit cast is `UnaryOp::Reinterpret` — the
    // same node Go, Java, C# and VB reach for `Float32bits` /
    // `floatToIntBits` / `BitConverter`. It used to be a call to a Fortran
    // profile row.
    let bitcast = |repr: vybe_ast::NumericRepr, arg: Expression| {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::Reinterpret(repr),
            expr: Box::new(arg),
        })
    };
    let source_is_real = source_hint.is_some_and(|hint| {
        let hint = hint.trim().to_ascii_lowercase();
        hint.starts_with("real") || hint.starts_with("double")
    }) || matches!(source.kind, ExprKind::Lit(Literal::Float(_)));
    let source_is_int = source_hint
        .is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("integer"))
        || matches!(source.kind, ExprKind::Lit(Literal::Int(_)));
    // Default REAL is 4 bytes, so `transfer(x, 0)` is the f32 pattern — but a
    // kind=8 / DOUBLE PRECISION operand has an 8-byte pattern and must not be
    // demoted to f32 on the way out.
    let is_kind8 = [source_hint, target_hint].iter().flatten().any(|hint| {
        let hint = hint.trim().to_ascii_lowercase();
        hint.contains("kind=8") || hint.contains("kind = 8") || hint.starts_with("double")
    });
    // The MOLD is only a type carrier, and it is just as often a NAME as a
    // literal: `sink = transfer(source, sink)`. When it is a name the assignment
    // target's declared type says the same thing.
    let hint_is_int = |hint: Option<&str>| {
        hint.is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("integer"))
    };
    let hint_is_real = |hint: Option<&str>| {
        hint.is_some_and(|hint| {
            let hint = hint.trim().to_ascii_lowercase();
            hint.starts_with("real") || hint.starts_with("double")
        })
    };
    let mold_is_int = matches!(mold.kind, ExprKind::Lit(Literal::Int(_)))
        || (matches!(mold.kind, ExprKind::Ident(_)) && hint_is_int(target_hint));
    let mold_is_real = matches!(mold.kind, ExprKind::Lit(Literal::Float(_)))
        || (matches!(mold.kind, ExprKind::Ident(_)) && hint_is_real(target_hint));
    if !fortran_type_hint_is_array(target_hint) {
        if source_is_real && mold_is_int {
            let name = if is_kind8 {
                vybe_ast::NumericRepr::I64
            } else {
                vybe_ast::NumericRepr::I32
            };
            return bitcast(name, source);
        }
        if source_is_int && mold_is_real {
            let name = if is_kind8 {
                vybe_ast::NumericRepr::F64
            } else {
                vybe_ast::NumericRepr::F32
            };
            return bitcast(name, source);
        }
    }
    if source_hint.is_some_and(is_fortran_string_type_hint) {
        if let Some(len) = source_hint.and_then(fortran_character_hint_len) {
            if len >= 2 {
                return build_fortran_char_bytes_to_int_expr(source, len);
            }
        }
        return build_fortran_char_code_expr(source);
    }
    if source_hint.is_some_and(fortran_type_hint_is_array_str) {
        if source_hint.is_some_and(fortran_type_hint_is_kind1_integer_array) {
            return build_fortran_byte_array_to_int_expr(source);
        }
        return build_fortran_first_array_item_expr(source);
    }
    if let Some(len) = fortran_string_literal_len(&mold) {
        return build_fortran_transfer_to_string_expr(source, len, size);
    }
    if let Some(value) = fortran_string_literal_to_int(&source) {
        return Expression::int(value);
    }
    match &source.kind {
        ExprKind::Lit(Literal::Bool(_)) => Expression::new(ExprKind::Ternary {
            cond: Box::new(source),
            then: Box::new(Expression::int(1)),
            else_: Box::new(Expression::int(0)),
        }),
        ExprKind::Array(_) => build_fortran_transfer_array_to_scalar_expr(source),
        _ => source,
    }
}

fn build_fortran_transfer_array_expr(
    source: Expression,
    size: Option<Expression>,
    source_hint: Option<&str>,
    target_hint: Option<&str>,
) -> Expression {
    if target_hint.is_some_and(fortran_type_hint_is_kind1_integer_array)
        && !source_hint.is_some_and(|hint| hint.contains("kind=1") || hint.contains("_1"))
    {
        return build_fortran_scalar_to_byte_array_expr(source, size);
    }
    // Only a literal array can be materialised element-by-element at walk time.
    // A NAME whose declared type is an array must keep its runtime value, or the
    // "pad and slice" path below is handed `[a]` — one element holding the whole
    // array, which is what made `target(1)` print `10,20,30,40`.
    // `transfer(s, 0, 1)` into an INTEGER array packs the character storage 4
    // bytes to the element, the same reading the scalar form does.
    if let (Some(len), Some(count)) = (
        source_hint
            .filter(|hint| is_fortran_string_type_hint(hint))
            .and_then(fortran_character_hint_len),
        size.as_ref().and_then(fortran_literal_int),
    ) {
        if target_hint.is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("integer"))
            && count > 0
        {
            let values = (0..count)
                .map(|index| {
                    build_fortran_char_bytes_to_int_at(source.clone(), len, index as usize * 4)
                })
                .collect::<Vec<_>>();
            return fortran_array_expr(values);
        }
    }
    let source_is_logical = source_hint
        .is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("logical"))
        || matches!(&source.kind, ExprKind::Array(items)
            if !items.is_empty()
                && items
                    .iter()
                    .all(|item| matches!(item.value.kind, ExprKind::Lit(Literal::Bool(_)))));
    if source_is_logical
        && target_hint.is_some_and(|hint| hint.trim().to_ascii_lowercase().starts_with("integer"))
        && size.is_none()
    {
        return build_fortran_logical_array_to_int_expr(source);
    }
    let source_is_runtime_array = !matches!(source.kind, ExprKind::Array(_))
        && source_hint.is_some_and(fortran_type_hint_is_array_str);
    if let Some(size_value) = size.as_ref().and_then(fortran_literal_int) {
        if source_is_runtime_array {
            // A NAME cannot be materialised element-by-element, but a LITERAL
            // size can be spelled as its own subscripts: `[a(1), a(2)]`.
            return build_fortran_transfer_sized_subscripts(&source, size_value);
        }
        return build_fortran_transfer_sized_array_literal(source, size_value);
    }
    let array_source = match &source.kind {
        ExprKind::Array(_) => source,
        _ if source_hint.is_some_and(is_fortran_complex_type_hint)
            || expr_is_fortran_complex_literalish(&source) =>
        {
            fortran_array_expr(vec![
            fortran_complex_real_part(&source),
            fortran_complex_imag_part(&source),
        ])
        }
        _ if source_hint.is_some_and(fortran_type_hint_is_rank_gt_one) => {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(source),
                    field: "flat".to_string(),
                    null_safe: false,
                })),
                args: Vec::new(),
                optional: false,
            })
        }
        ExprKind::Object(_) | ExprKind::New { .. } => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Object")),
                field: "values".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(source)],
            optional: false,
        }),
        // A rank-1 array-typed name already IS the element sequence.
        _ if source_hint.is_some_and(fortran_type_hint_is_array_str) => source,
        _ => fortran_array_expr(vec![source]),
    };
    let Some(size) = size else {
        return array_source;
    };
    let padded = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_source.clone()),
            field: "concat".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![Argument::positional(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("max")),
                        args: vec![
                            Argument::positional(Expression::int(0)),
                            Argument::positional(Expression::new(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(size.clone()),
                                right: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(array_source),
                                    field: "length".to_string(),
                                    null_safe: false,
                                })),
                            })),
                        ],
                        optional: false,
                    }))],
                    optional: false,
                })),
                field: "fill".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::int(0))],
            optional: false,
        }))],
        optional: false,
    });
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(padded),
            field: "slice".to_string(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(Expression::int(0)),
            Argument::positional(size),
        ],
        optional: false,
    })
}

/// `transfer(a, mold, n)` where `a` is an array-typed NAME: the first `n`
/// elements, spelled as Fortran's own 1-based subscripts.
fn build_fortran_transfer_sized_subscripts(source: &Expression, size: i64) -> Expression {
    if size <= 0 {
        return fortran_array_expr(Vec::new());
    }
    let values = (1..=size)
        .map(|index| {
            Expression::new(ExprKind::Index {
                object: Box::new(source.clone()),
                index: Box::new(Expression::int(index)),
                null_safe: false,
            })
        })
        .collect::<Vec<_>>();
    fortran_array_expr(values)
}

fn build_fortran_transfer_sized_array_literal(source: Expression, size: i64) -> Expression {
    if size <= 0 {
        return fortran_array_expr(Vec::new());
    }
    let mut values = Vec::with_capacity(size as usize);
    match source.kind {
        ExprKind::Array(items) => {
            for item in items.into_iter().take(size as usize) {
                values.push(item.value);
            }
        }
        _ => values.push(source),
    }
    while values.len() < size as usize {
        values.push(Expression::int(0));
    }
    fortran_array_expr(values)
}

fn build_fortran_scalar_to_byte_array_expr(
    source: Expression,
    size: Option<Expression>,
) -> Expression {
    let n = fortran_literal_int(&size.unwrap_or_else(|| Expression::int(8))).unwrap_or(8);
    let values = (0..n.max(0) as usize)
        .map(|idx| {
            let shifted = if idx == 0 {
                source.clone()
            } else {
                Expression::new(ExprKind::Binary {
                    op: BinOp::Shr,
                    left: Box::new(source.clone()),
                    right: Box::new(Expression::int((idx * 8) as i64)),
                })
            };
            Expression::new(ExprKind::Binary {
                op: BinOp::BitAnd,
                left: Box::new(shifted),
                right: Box::new(Expression::int(255)),
            })
        })
        .collect();
    fortran_array_expr(values)
}

fn build_fortran_transfer_to_string_expr(
    source: Expression,
    len: usize,
    size: Option<Expression>,
) -> Expression {
    let n = fortran_literal_int(&size.unwrap_or_else(|| Expression::int(len as i64)))
        .unwrap_or(len as i64)
        .max(1) as usize;
    if n == 1 {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("char")),
            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                op: BinOp::BitAnd,
                left: Box::new(source),
                right: Box::new(Expression::int(255)),
            }))],
            optional: false,
        });
    }
    let chars = (0..n)
        .map(|idx| {
            let shifted = if idx == 0 {
                source.clone()
            } else {
                Expression::new(ExprKind::Binary {
                    op: BinOp::Shr,
                    left: Box::new(source.clone()),
                    right: Box::new(Expression::int((idx * 8) as i64)),
                })
            };
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("char")),
                args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(shifted),
                    right: Box::new(Expression::int(255)),
                }))],
                optional: false,
            })
        })
        .collect::<Vec<_>>();
    concat_fortran_io_parts(chars)
}

fn build_fortran_transfer_array_to_scalar_expr(source: Expression) -> Expression {
    let ExprKind::Array(items) = &source.kind else {
        return source;
    };
    if items.is_empty() {
        return Expression::int(0);
    }
    if items
        .iter()
        .all(|item| matches!(item.value.kind, ExprKind::Lit(Literal::Bool(_))))
    {
        return Expression::new(ExprKind::Ternary {
            cond: Box::new(items[0].value.clone()),
            then: Box::new(Expression::int(1)),
            else_: Box::new(Expression::int(0)),
        });
    }
    items[0].value.clone()
}

fn build_fortran_byte_array_to_int_expr(source: Expression) -> Expression {
    let mut result = Expression::int(0);
    for idx in 0..8 {
        // A Fortran subscript is 1-based, and subscript 0 wrapped to the LAST
        // element — which is why `[18,52,86,120]` packed to 0x56341278 instead
        // of 0x78563412.
        let item = Expression::new(ExprKind::Index {
            object: Box::new(source.clone()),
            index: Box::new(Expression::int(idx + 1)),
            null_safe: false,
        });
        let byte = Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(item),
            right: Box::new(Expression::int(255)),
        });
        let shifted = if idx == 0 {
            byte
        } else {
            Expression::new(ExprKind::Binary {
                op: BinOp::Shl,
                left: Box::new(byte),
                right: Box::new(Expression::int((idx * 8) as i64)),
            })
        };
        result = Expression::new(ExprKind::Binary {
            op: BinOp::BitOr,
            left: Box::new(result),
            right: Box::new(shifted),
        });
    }
    result
}

fn build_fortran_first_array_item_expr(source: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(source),
        // A Fortran subscript is 1-based; subscript 0 wraps to the LAST element.
        index: Box::new(Expression::int(1)),
        null_safe: false,
    })
}

fn build_fortran_complex_expr_from_array(source: Expression) -> Expression {
    build_fortran_complex_expr(
        Expression::new(ExprKind::Index {
            object: Box::new(source.clone()),
            index: Box::new(Expression::int(0)),
            null_safe: false,
        }),
        Expression::new(ExprKind::Index {
            object: Box::new(source),
            index: Box::new(Expression::int(1)),
            null_safe: false,
        }),
    )
}

/// The declared length of a CHARACTER type hint (`character(len=2)`,
/// `character*4`), when it is a literal count.
fn fortran_character_hint_len(type_hint: &str) -> Option<usize> {
    let lower = type_hint.trim().to_ascii_lowercase();
    let rest = lower.strip_prefix("character")?.trim_start();
    let digits: String = if let Some(rest) = rest.strip_prefix('*') {
        rest.trim_start()
            .chars()
            .take_while(|c| c.is_ascii_digit())
            .collect()
    } else {
        let inner = rest.strip_prefix('(')?.trim_end_matches(')').trim();
        let inner = inner
            .strip_prefix("len")
            .map(|rest| rest.trim_start().trim_start_matches('=').trim_start())
            .unwrap_or(inner);
        inner.chars().take_while(|c| c.is_ascii_digit()).collect()
    };
    digits.parse().ok()
}

/// `transfer('AB', 0)` is 16961 — the two bytes read little-endian, not the
/// code of the first character. A 4-byte integer holds at most 4 of them, and
/// storage past the declared length reads as zero.
/// `transfer(mask, n)` from LOGICAL to INTEGER: `.true.` stores as 1, not as a
/// boolean. The elements are the storage, so each one is coerced in place.
fn build_fortran_logical_array_to_int_expr(source: Expression) -> Expression {
    let param = "__fortran_transfer_bit";
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(source),
            field: "map".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::ident(param)),
                then: Box::new(Expression::int(1)),
                else_: Box::new(Expression::int(0)),
            }))),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    })
}

fn build_fortran_char_bytes_to_int_expr(source: Expression, len: usize) -> Expression {
    build_fortran_char_bytes_to_int_at(source, len, 0)
}

/// The integer holding the 4 bytes starting at `offset`; storage past `len`
/// reads as zero.
fn build_fortran_char_bytes_to_int_at(
    source: Expression,
    len: usize,
    offset: usize,
) -> Expression {
    let mut result = Expression::int(0);
    for idx in 0..len.saturating_sub(offset).min(4) {
        let code = build_fortran_char_code_at_expr(source.clone(), (offset + idx) as i64);
        let shifted = if idx == 0 {
            code
        } else {
            Expression::new(ExprKind::Binary {
                op: BinOp::Shl,
                left: Box::new(code),
                right: Box::new(Expression::int((idx * 8) as i64)),
            })
        };
        result = Expression::new(ExprKind::Binary {
            op: BinOp::BitOr,
            left: Box::new(result),
            right: Box::new(shifted),
        });
    }
    result
}

fn build_fortran_char_code_at_expr(source: Expression, index: i64) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(source),
            field: "charCodeAt".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::int(index))],
        optional: false,
    })
}

fn build_fortran_char_code_expr(source: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(source),
            field: "charCodeAt".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::int(0))],
        optional: false,
    })
}

fn fortran_array_expr(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| vybe_ast::ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn fortran_string_literal_len(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.len()),
        _ => None,
    }
}

fn fortran_string_literal_to_int(expr: &Expression) -> Option<i64> {
    let ExprKind::Lit(Literal::Str(value)) = &expr.kind else {
        return None;
    };
    let mut result = 0_i64;
    for (idx, byte) in value.bytes().take(8).enumerate() {
        result |= (byte as i64) << (idx * 8);
    }
    Some(result)
}

fn expr_is_fortran_complex_literalish(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Object(_))
}

fn fortran_type_hint_is_array_str(type_hint: &str) -> bool {
    type_hint.contains("()")
}

fn fortran_type_hint_is_rank_gt_one(type_hint: &str) -> bool {
    type_hint.matches("()").count() > 1
}

fn fortran_type_hint_is_kind1_integer_array(type_hint: &str) -> bool {
    let lower = type_hint.to_ascii_lowercase();
    lower.starts_with("integer") && lower.contains("kind=1") && lower.contains("()")
}

/// Every element of a ranked array as one flat run.
///
/// `ALL(mask)`, `ANY(mask)` and `COUNT(mask)` with no `dim` ask about EVERY
/// element whatever the rank, and a rank-2 array is a NEST: mapping over it
/// visits rows, a row is an array, and an array is always truthy. So `any`
/// answered true for an all-false mask and `count` counted rows. `sum` never
/// had the problem because it reaches a rank-aware lowering; these three do
/// not, and flattening is the rank-independent answer — on a rank-1 array it is
/// a no-op.
fn build_fortran_flatten_expr(array: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array),
            field: "flat".to_string(),
            null_safe: false,
        })),
        // Fortran caps rank at 15; nothing nests deeper than that.
        args: vec![Argument::positional(Expression::int(15))],
        optional: false,
    })
}

fn build_fortran_logical_array_reducer(array_expr: Expression, method: &str) -> Expression {
    let item_name = "__fortran_logical_item";
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(build_fortran_flatten_expr(array_expr)),
            field: method.to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: item_name.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(fortran_expr_is_true(Expression::ident(item_name)))),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    })
}

fn build_fortran_count_expr(array_expr: Expression) -> Expression {
    let item_name = "__fortran_count_item";
    let source = build_fortran_flatten_expr(build_fortran_count_source_expr(array_expr));
    let counted = build_fortran_array_map(
        source,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(fortran_expr_is_true(Expression::ident(item_name))),
            then: Box::new(Expression::int(1)),
            else_: Box::new(Expression::int(0)),
        }),
        false,
        item_name,
        "__fortran_count_index",
    );
    build_fortran_array_reduction("sum", counted, 0)
}

fn build_fortran_count_source_expr(expr: Expression) -> Expression {
    let ExprKind::Binary { op, left, right } = &expr.kind else {
        return expr;
    };
    let left = (**left).clone();
    let right = (**right).clone();
    let item_name = "__fortran_count_pred_item";
    let index_name = "__fortran_count_pred_index";

    // `count` builds its own comparison rather than reading the one the
    // elementwise lowering would have produced, so it has to flatten its
    // operands ITSELF — mapping a rank-2 operand compares whole ROWS, and
    // `row == 1` is false for every row. Two ranked operands have the same
    // shape, so flattening both keeps them element-for-element aligned. The
    // test is taken BEFORE wrapping: a wrapped operand is a method call, which
    // is not a spelling `fortran_count_arrayish_expr` recognises.
    let left_is_array = fortran_count_arrayish_expr(&left);
    let right_is_array = fortran_count_arrayish_expr(&right);
    let left = if left_is_array {
        build_fortran_flatten_expr(left)
    } else {
        left
    };
    let right = if right_is_array {
        build_fortran_flatten_expr(right)
    } else {
        right
    };

    if left_is_array && right_is_array {
        return build_fortran_array_map(
            left,
            Expression::new(ExprKind::Binary {
                op: *op,
                left: Box::new(Expression::ident(item_name)),
                right: Box::new(Expression::new(ExprKind::Index {
                    object: Box::new(right),
                    index: Box::new(Expression::ident(index_name)),
                    null_safe: false,
                })),
            }),
            true,
            item_name,
            index_name,
        );
    }

    if left_is_array {
        return build_fortran_array_map(
            left,
            Expression::new(ExprKind::Binary {
                op: *op,
                left: Box::new(Expression::ident(item_name)),
                right: Box::new(right),
            }),
            false,
            item_name,
            index_name,
        );
    }

    if right_is_array {
        return build_fortran_array_map(
            right,
            Expression::new(ExprKind::Binary {
                op: *op,
                left: Box::new(left),
                right: Box::new(Expression::ident(item_name)),
            }),
            false,
            item_name,
            index_name,
        );
    }

    expr
}

fn fortran_count_arrayish_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(_)
        | ExprKind::Member { .. }
        | ExprKind::Array(_)
        | ExprKind::Slice { .. }
        | ExprKind::ArrayMap { .. }
        | ExprKind::ArrayTransform { .. } => true,
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => name.eq_ignore_ascii_case("Array"),
            ExprKind::Member { field, .. } => matches!(
                field.to_ascii_lowercase().as_str(),
                "map" | "filter" | "flatmap"
            ),
            _ => false,
        },
        _ => false,
    }
}

fn build_fortran_lexical_compare_expr(
    op: BinOp,
    left: Expression,
    right: Expression,
) -> Expression {
    let compare = Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(compare),
        then: Box::new(Expression::bool(true)),
        else_: Box::new(Expression::bool(false)),
    })
}

fn build_fortran_dot_product_expr(left: Expression, right: Expression) -> Expression {
    let left_item_name = "__fortran_dot_left";
    let left_index_name = "__fortran_dot_index";

    let product = Expression::new(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(Expression::ident(left_item_name)),
        right: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(right),
            index: Box::new(Expression::ident(left_index_name)),
            null_safe: false,
        })),
    });

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("sum")),
        args: vec![Argument::positional(build_fortran_array_map(
            left,
            product,
            true,
            left_item_name,
            left_index_name,
        ))],
        optional: false,
    })
}

fn fortran_expr_is_literal_bool(expr: &Expression) -> Option<bool> {
    match &expr.kind {
        ExprKind::Lit(Literal::Bool(value)) => Some(*value),
        ExprKind::Ident(name) if name.eq_ignore_ascii_case(".true.") => Some(true),
        ExprKind::Ident(name) if name.eq_ignore_ascii_case(".false.") => Some(false),
        _ => None,
    }
}

fn fortran_literal_int(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) => Some(*value as i64),
        _ => None,
    }
}

/// An optional intrinsic argument, by keyword or by position.
///
/// Fortran lets every argument be written either way, and a keyword one is not
/// in the positional list at all — so reading only positions drops
/// `pad=[0]` silently, and reading only keywords drops `reshape(a, s, [0])`.
fn fortran_argument(
    keyword: &str,
    position: usize,
    args: &[Argument],
    positional_args: &[Expression],
) -> Option<Expression> {
    args.iter()
        .find(|arg| {
            arg.name
                .as_deref()
                .is_some_and(|name| name.eq_ignore_ascii_case(keyword))
        })
        .map(|arg| arg.value.clone())
        .or_else(|| positional_args.get(position).cloned())
}

/// `ORDER=` — a permutation of `1..rank` saying which subscript varies fastest.
///
/// The identity permutation is Fortran's own element order, the full reversal
/// is C's, and those are the two the shared traversal order names. Any other
/// permutation is a genuine reordering this cannot express, and answering it
/// with one of the two would be a wrong answer rather than a missing one — so
/// it declines to fold and the call is left to fail loudly.
fn fortran_traversal_order(order: &Expression) -> Option<ArrayTraversalOrder> {
    let ExprKind::Array(items) = &order.kind else {
        return None;
    };
    let dims: Vec<i64> = items
        .iter()
        .map(|item| fortran_literal_int(&item.value))
        .collect::<Option<_>>()?;
    let ascending: Vec<i64> = (1..=dims.len() as i64).collect();
    if dims == ascending {
        return Some(ArrayTraversalOrder::ColumnMajor);
    }
    if dims.iter().rev().copied().eq(ascending) {
        return Some(ArrayTraversalOrder::RowMajor);
    }
    None
}

fn build_fortran_array_length_expr(array: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(array),
        field: "length".to_string(),
        null_safe: false,
    })
}

fn build_fortran_normalized_circular_shift(shift: Expression, size: Expression) -> Expression {
    let mod_shift = build_fortran_modulo_expr(shift, size.clone());
    build_fortran_modulo_expr(
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(mod_shift),
            right: Box::new(size.clone()),
        }),
        size,
    )
}

fn build_fortran_cshift_1d_expr(array: Expression, shift: Expression) -> Expression {
    let size = build_fortran_array_length_expr(array.clone());
    let effective_shift = build_fortran_normalized_circular_shift(shift, size.clone());
    let index_name = "__fortran_cshift_index";
    let item_name = "__fortran_cshift_item";
    let source_index = Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(build_fortran_modulo_expr(
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(Expression::ident(index_name)),
                    right: Box::new(Expression::int(1)),
                })),
                right: Box::new(effective_shift),
            }),
            size.clone(),
        )),
        right: Box::new(Expression::int(1)),
    });
    let body = Expression::new(ExprKind::Index {
        object: Box::new(array),
        index: Box::new(source_index),
        null_safe: false,
    });
    build_fortran_array_map(
        build_fortran_array_fill(size, Expression::int(0)),
        body,
        true,
        item_name,
        index_name,
    )
}

fn build_fortran_eoshift_1d_expr(
    array: Expression,
    shift: Expression,
    boundary: Expression,
) -> Expression {
    let size = build_fortran_array_length_expr(array.clone());
    let index_name = "__fortran_eoshift_index";
    let item_name = "__fortran_eoshift_item";
    let shifted_index = Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(Expression::ident(index_name)),
        right: Box::new(shift.clone()),
    });
    let in_range = Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(shift),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::LtEq,
            left: Box::new(shifted_index.clone()),
            right: Box::new(size.clone()),
        })),
        else_: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(shifted_index.clone()),
            right: Box::new(Expression::int(1)),
        })),
    });
    let body = Expression::new(ExprKind::Ternary {
        cond: Box::new(in_range),
        then: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(array),
            index: Box::new(shifted_index),
            null_safe: false,
        })),
        else_: Box::new(boundary),
    });
    build_fortran_array_map(
        build_fortran_array_fill(size, Expression::int(0)),
        body,
        true,
        item_name,
        index_name,
    )
}

fn build_fortran_spread_dim1_expr(source: Expression, ncopies: Expression) -> Expression {
    build_fortran_array_fill(ncopies, source)
}

fn build_fortran_spread_dim2_expr(source: Expression, ncopies: Expression) -> Expression {
    let item_name = "__fortran_spread_dim2_item";
    let index_name = "__fortran_spread_dim2_index";
    let row = build_fortran_array_fill(ncopies, Expression::ident(item_name));
    build_fortran_array_map(source, row, false, item_name, index_name)
}

/// "Did this argument arrive as a whole array, or as one value?" — the question
/// `merge` has to ask of its mask and of each source.
///

/// `MERGE(tsource, fsource, mask)` as the shared ranked-array node.
///
/// Column-major because that is Fortran's element order, the same choice
/// `PACK`/`UNPACK`/`RESHAPE` already make here.
fn fortran_merge_node(
    true_source: Expression,
    false_source: Expression,
    mask: Expression,
) -> Expression {
    Expression::new(ExprKind::ArrayTransform {
        op: ArrayTransformOp::MergeMask,
        args: vec![true_source, false_source, mask],
        order: ArrayTraversalOrder::ColumnMajor,
    })
}

fn build_fortran_transpose_expr(matrix: Expression) -> Expression {
    let column_item_name = "__fortran_transpose_column_item";
    let column_index_name = "__fortran_transpose_column_index";
    let row_item_name = "__fortran_transpose_row";

    let first_row = Expression::new(ExprKind::Index {
        object: Box::new(matrix.clone()),
        index: Box::new(Expression::int(1)),
        null_safe: false,
    });

    build_fortran_array_map(
        first_row,
        build_fortran_array_map(
            matrix,
            Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(row_item_name)),
                index: Box::new(Expression::ident(column_index_name)),
                null_safe: false,
            }),
            false,
            row_item_name,
            "__fortran_transpose_row_index",
        ),
        true,
        column_item_name,
        column_index_name,
    )
}

/// `ishft(i, shift)` — shift left, or right when `shift` is negative.
///
/// The right shift is LOGICAL. `ishft` moves a BIT PATTERN; it has no notion of
/// a sign to preserve, so `ishft(-2, -1)` is 2147483647 and not −1. An
/// arithmetic `Shr` here was answering with the sign smeared back in.
fn build_fortran_ishft_expr(value: Expression, shift: Expression) -> Expression {
    let negated_shift = Expression::new(ExprKind::Unary {
        op: UnaryOp::Neg,
        expr: Box::new(shift.clone()),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(shift.clone()),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Shl,
            left: Box::new(value.clone()),
            right: Box::new(shift),
        })),
        else_: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::UShr,
            left: Box::new(value),
            right: Box::new(negated_shift),
        })),
    })
}

fn build_fortran_huge_expr(arg: &Expression) -> Expression {
    match &arg.kind {
        // Default `real` is kind 4, so `huge(1.0)` is the largest FLOAT, not the
        // largest double — `huge(1.0d0)` would be the other one.
        ExprKind::Lit(Literal::Float(_)) => Expression::float(f32::MAX as f64),
        _ => Expression::int(i32::MAX as i64),
    }
}

#[derive(Clone, Copy)]
struct FortranInquiryModel {
    /// The KIND number — 4 for a default `integer`/`real`/`logical`, 1 for a
    /// `character`. Not derivable from `bits`: a default real is kind 4 and 32
    /// bits, but a kind-1 character is 8, and `kind` was previously answered by
    /// a hard-coded `8` in the shared compiler for every type alike.
    kind: i64,
    bits: i64,
    precision: Option<i64>,
    range: i64,
    digits: i64,
}

/// The gfortran model for an integer of `kind`, as `(bits, range, digits)`.
///
/// Measured, not derived — `digits` happens to be `bits - 1` across the board
/// but `range` does not follow any formula worth writing down.
fn fortran_integer_kind_model(kind: i64) -> (i64, i64, i64) {
    match kind {
        1 => (8, 2, 7),
        2 => (16, 4, 15),
        8 => (64, 18, 63),
        16 => (128, 38, 127),
        _ => (32, 9, 31),
    }
}

/// The gfortran model for a real of `kind`, as `(bits, precision, range, digits)`.
fn fortran_real_kind_model(kind: i64) -> (i64, i64, i64, i64) {
    match kind {
        8 => (64, 15, 307, 53),
        16 => (128, 33, 4931, 113),
        _ => (32, 6, 37, 24),
    }
}

/// The KIND a numeric LITERAL spells, read from its source text.
///
/// Only answerable here: `1.0d0`, `1.0_8` and `1.0` all become
/// `Literal::Float`, so the AST cannot distinguish them. Returns `None` for
/// anything that is not a literal, leaving the ordinary type-inquiry fold to
/// answer from the declared type.
/// `OUT_OF_RANGE(X, MOLD [, ROUND])` — F2018 16.9.146. True when X cannot be
/// represented in MOLD's type and kind.
///
/// MOLD is a value used only for its TYPE AND KIND, so the range comes from the
/// literal's kind suffix: `0_1` is int8's [-128, 127], `0` is default integer.
/// That suffix is only readable from the SOURCE TEXT, which is why this folds
/// where `arg_texts` is still in scope rather than in the intrinsic pass.
///
/// A real X is converted before the test — truncated toward zero, or rounded
/// when ROUND is true. NaN needs its own arm because every comparison against
/// it is false; ±Infinity needs none, since it fails the bounds test already.
fn build_fortran_out_of_range_expr(
    value: Expression,
    mold_text: &str,
    round: Option<Expression>,
) -> Option<Expression> {
    let lowered = mold_text.trim().to_ascii_lowercase();
    let kind = fortran_literal_kind_from_text(&lowered).unwrap_or(4);
    let mold_is_real = lowered.contains('.') || lowered.contains('e') || lowered.contains('d');

    let (lo, hi) = if mold_is_real {
        match kind {
            8 => (f64::MIN, f64::MAX),
            _ => (-(f32::MAX as f64), f32::MAX as f64),
        }
    } else {
        let (l, h) = match kind {
            1 => (i8::MIN as i64, i8::MAX as i64),
            2 => (i16::MIN as i64, i16::MAX as i64),
            8 => (i64::MIN, i64::MAX),
            _ => (i32::MIN as i64, i32::MAX as i64),
        };
        (l as f64, h as f64)
    };

    // Truncation toward zero is `int`; `nint` rounds. Both are already lowered.
    let converted = if mold_is_real {
        value.clone()
    } else {
        match round {
            Some(flag) => fortran_ternary(
                fortran_expr_is_true(flag),
                fortran_call("nint", vec![value.clone()]),
                fortran_call("int", vec![value.clone()]),
            ),
            None => fortran_call("int", vec![value.clone()]),
        }
    };

    let num = |v: f64| {
        if v.fract() == 0.0 && v.abs() < 9.0e15 {
            Expression::int(v as i64)
        } else {
            Expression::float(v)
        }
    };
    let is_nan = fortran_bin(BinOp::NotEq, value.clone(), value);
    let below = fortran_bin(BinOp::Lt, converted.clone(), num(lo));
    let above = fortran_bin(BinOp::Gt, converted, num(hi));
    Some(fortran_bin(
        BinOp::Or,
        is_nan,
        fortran_bin(BinOp::Or, below, above),
    ))
}

fn fortran_literal_kind_from_text(text: &str) -> Option<i64> {
    let text = text.trim();
    // A complex literal `(1.0, 2.0)` — kind comes from its components, and a
    // complex of kind k is two reals of kind k.
    if let Some(inner) = text.strip_prefix('(').and_then(|t| t.strip_suffix(')')) {
        if let Some((real, imaginary)) = inner.split_once(',') {
            let real_kind = fortran_literal_kind_from_text(real)?;
            let imaginary_kind = fortran_literal_kind_from_text(imaginary)?;
            return Some(real_kind.max(imaginary_kind));
        }
    }
    let lowered = text.to_ascii_lowercase();
    if !lowered.starts_with(|c: char| c.is_ascii_digit() || c == '.' || c == '-' || c == '+') {
        return None;
    }
    // `1.0_8` / `1_2` — an explicit kind suffix outranks everything.
    if let Some((_, suffix)) = lowered.rsplit_once('_') {
        if let Ok(kind) = suffix.parse::<i64>() {
            return matches!(kind, 1 | 2 | 4 | 8 | 16).then_some(kind);
        }
        // `1.0_dp` names a kind PARAMETER whose value is not known here.
        return None;
    }
    if !lowered.chars().all(|c| {
        c.is_ascii_digit() || matches!(c, '.' | '+' | '-' | 'e' | 'd')
    }) {
        return None;
    }
    // `1.0d0` is DOUBLE precision — that `d` is the whole difference.
    if lowered.contains('d') {
        return Some(8);
    }
    if lowered.contains('.') || lowered.contains('e') {
        return Some(4);
    }
    Some(4)
}

/// The value of a compile-time integer argument, if it is one.
///
/// `selected_int_kind` / `selected_real_kind` take a *constant* expression in
/// every real program — the result names a TYPE — so folding is the whole
/// implementation, and a non-constant argument correctly folds to nothing
/// rather than to a wrong guess.
fn fortran_const_int(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) => Some(*value as i64),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => fortran_const_int(expr).map(|value| -value),
        ExprKind::Unary {
            op: UnaryOp::Pos,
            expr,
        } => fortran_const_int(expr),
        _ => None,
    }
}

fn fortran_integer_model(kind: i64) -> FortranInquiryModel {
    let (bits, range, digits) = fortran_integer_kind_model(kind);
    FortranInquiryModel {
        kind,
        bits,
        precision: None,
        range,
        digits,
    }
}

fn fortran_real_model(kind: i64) -> FortranInquiryModel {
    let (bits, precision, range, digits) = fortran_real_kind_model(kind);
    FortranInquiryModel {
        kind,
        bits,
        precision: Some(precision),
        range,
        digits,
    }
}

fn fortran_logical_model() -> FortranInquiryModel {
    FortranInquiryModel {
        kind: 4,
        bits: 32,
        precision: None,
        range: 1,
        digits: 1,
    }
}

/// `selected_int_kind(r)` — the smallest integer kind whose decimal exponent
/// range reaches `r`, or −1 when none does.
fn fortran_selected_int_kind(requested_range: i64) -> i64 {
    [1i64, 2, 4, 8, 16]
        .into_iter()
        .find(|kind| fortran_integer_kind_model(*kind).1 >= requested_range)
        .unwrap_or(-1)
}

/// `selected_real_kind(p, r)` — the smallest real kind meeting BOTH the
/// requested decimal precision and range, or −1 when none does.
fn fortran_selected_real_kind(precision: i64, range: i64) -> i64 {
    [4i64, 8, 16]
        .into_iter()
        .find(|kind| {
            let (_, kind_precision, kind_range, _) = fortran_real_kind_model(*kind);
            kind_precision >= precision && kind_range >= range
        })
        .unwrap_or(-1)
}

fn fortran_inquiry_model_from_hint(type_hint: &str) -> Option<FortranInquiryModel> {
    let t = type_hint.to_ascii_lowercase();
    let spelled_kind = fortran_spelled_kind(&t);
    if t.contains("integer") {
        // DEFAULT integer is kind 4, not 8. `digits` and `range` were also
        // being set to the same number, which is only ever right by accident.
        let kind = spelled_kind.unwrap_or(4);
        let (bits, range, digits) = fortran_integer_kind_model(kind);
        return Some(FortranInquiryModel {
            kind,
            bits,
            precision: None,
            range,
            digits,
        });
    }
    if t.contains("real") || t.contains("double precision") {
        // `double precision` IS kind 8; a bare `real` is kind 4. The old code
        // had it backwards — anything not spelled `kind=4` was treated as
        // double — and it had `precision` and `digits` SWAPPED: for a default
        // real gfortran gives precision 6 and digits 24, not 24 and 6.
        let kind = spelled_kind.unwrap_or(if t.contains("double precision") { 8 } else { 4 });
        let (bits, precision, range, digits) = fortran_real_kind_model(kind);
        return Some(FortranInquiryModel {
            kind,
            bits,
            precision: Some(precision),
            range,
            digits,
        });
    }
    if t.contains("complex") {
        // A complex is two reals of its kind, so `kind`/`precision`/`range`
        // are the real model's and only the storage doubles.
        let kind = spelled_kind.unwrap_or(if t.contains("double complex") { 8 } else { 4 });
        let (bits, precision, range, digits) = fortran_real_kind_model(kind);
        return Some(FortranInquiryModel {
            kind,
            bits: bits * 2,
            precision: Some(precision),
            range,
            digits,
        });
    }
    if t.contains("logical") {
        return Some(FortranInquiryModel {
            kind: spelled_kind.unwrap_or(4),
            bits: spelled_kind.unwrap_or(4) * 8,
            precision: None,
            range: 1,
            digits: 1,
        });
    }
    if t.contains("character") {
        return Some(FortranInquiryModel {
            kind: 1,
            bits: 8,
            precision: None,
            range: 0,
            digits: 0,
        });
    }
    None
}

/// The KIND a declaration spells out, if it spells one — `integer(kind=8)`,
/// `real*8`, `integer(2)`.
///
/// Reads the NUMBER rather than testing for one spelling at a time: the
/// previous code asked `contains("kind=1") || contains("*1")` per kind and had
/// no arm for `(4)` on an integer or `kind=16` on anything.
fn fortran_spelled_kind(lowered_hint: &str) -> Option<i64> {
    for marker in ["kind=", "*", "("] {
        let Some(at) = lowered_hint.find(marker) else {
            continue;
        };
        let rest = &lowered_hint[at + marker.len()..];
        let digits: String = rest.chars().take_while(char::is_ascii_digit).collect();
        if let Ok(kind) = digits.parse::<i64>() {
            // `character(len=8)` is a LENGTH, not a kind — and a character is
            // kind 1 regardless, so the caller ignores this anyway.
            if matches!(kind, 1 | 2 | 4 | 8 | 16) {
                return Some(kind);
            }
        }
    }
    None
}

/// Is this expression a kind-8 (double precision) real?
///
/// Default `real` is kind 4, so f32 is the answer whenever nothing says
/// otherwise — `SPACING(1.0)` is 2^-23, and answering 2^-52 is wrong by 29
/// binades.
fn fortran_expr_is_kind8_real(expr: &Expression, type_env: &HashMap<String, String>) -> bool {
    fortran_inquiry_model_from_expr(expr, type_env).is_some_and(|model| model.kind == 8)
}

fn fortran_inquiry_model_from_expr(
    expr: &Expression,
    type_env: &HashMap<String, String>,
) -> Option<FortranInquiryModel> {
    match &expr.kind {
        // `kind(1.0)` is 4 — an unsuffixed real literal is DEFAULT real, not
        // double. It was modelled as kind 8 here, which is what made every
        // `precision`/`range`/`digits` answer double-precision numbers.
        ExprKind::Lit(Literal::Float(_)) => Some(fortran_real_model(4)),
        ExprKind::Lit(Literal::Int(_)) => Some(fortran_integer_model(4)),
        ExprKind::Lit(Literal::Bool(_)) => Some(fortran_logical_model()),
        ExprKind::Lit(Literal::Str(value)) => Some(FortranInquiryModel {
            kind: 1,
            // `storage_size` of a character string is its LENGTH in bits.
            bits: (value.chars().count().max(1) as i64) * 8,
            precision: None,
            range: 0,
            digits: 0,
        }),
        ExprKind::Lit(Literal::Char(_)) => Some(FortranInquiryModel {
            kind: 1,
            bits: 8,
            precision: None,
            range: 0,
            digits: 0,
        }),
        ExprKind::Ident(name)
            if name.eq_ignore_ascii_case(".true.") || name.eq_ignore_ascii_case(".false.") =>
        {
            Some(fortran_logical_model())
        }
        ExprKind::Ident(name) => type_env
            .get(&name.to_ascii_lowercase())
            .and_then(|hint| fortran_inquiry_model_from_hint(hint)),
        _ => None,
    }
}

fn fold_fortran_type_inquiry(
    name: &str,
    arg: &Expression,
    type_env: &HashMap<String, String>,
    _kind_arg: Option<&Expression>,
) -> Option<Expression> {
    let model = fortran_inquiry_model_from_expr(arg, type_env)?;
    let value = match name {
        // `kind` was answered by a hard-coded `I32(8)` in the SHARED compiler
        // (`primitives/builtins.rs`), for every type and every language that
        // spells a function `kind`. Folding it here means the constant never
        // reaches that arm.
        "kind" => model.kind,
        "bit_size" | "storage_size" => model.bits,
        "precision" => model.precision?,
        "range" => model.range,
        "digits" => model.digits,
        // The binary exponent range, IEEE-754 for each width. Only a REAL has
        // one — gfortran rejects `maxexponent` of an integer — and `precision`
        // is `Some` exactly for reals, so the model already carries the
        // discriminator and needs no new field.
        "maxexponent" | "minexponent" => {
            model.precision?;
            let (max, min) = match model.bits {
                64 => (1024, -1021),
                128 => (16384, -16381),
                _ => (128, -125),
            };
            if name == "maxexponent" {
                max
            } else {
                min
            }
        }
        _ => return None,
    };
    Some(Expression::int(value))
}

fn lower_fortran_type_inquiry_in_expr(expr: &mut Expression, type_env: &HashMap<String, String>) {
    match &mut expr.kind {
        // `MERGE`/`PACK`/`UNPACK`/`RESHAPE` are ArrayTransform NODES, and
        // their arguments are ORDINARY expressions. Without an arm here the
        // pass walks straight past them — which is how `nearest(...)` inside
        // a `merge(...)` stopped being folded the moment MERGE became a node
        // instead of a call.
        ExprKind::ArrayTransform { args, .. } => {
            for arg in args.iter_mut() {
                lower_fortran_type_inquiry_in_expr(arg, type_env);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                lower_fortran_type_inquiry_in_expr(&mut arg.value, type_env);
            }
            let ExprKind::Ident(name) = &callee.kind else {
                return;
            };
            let lowered = name.to_ascii_lowercase();
            // Collected BEFORE any `*expr` assignment below: the borrow of
            // `expr.kind` that produced `args` has to end first.
            //
            // `SPACING` and `NEAREST` are numeric INQUIRY functions in
            // Fortran's own classification, and like `kind` they need the
            // declared type: default `real` is kind 4 (ULP 2^-23), `double
            // precision` is kind 8 (2^-52). Answering in one lane is wrong for
            // the other half the time.
            //
            // ⛔ `NEAREST(X, S)` reads only the SIGN of S — it is a DIRECTION,
            // not a target. Lowering it to `nextafter(x, s)` walks TOWARD `s`
            // and gives the opposite neighbour whenever `s` is on the far side:
            // `nearest(1000.0d0, 1.0d0)` steps UP, not down.
            let positional = args
                .iter()
                .filter(|arg| arg.name.is_none())
                .map(|arg| arg.value.clone())
                .collect::<Vec<_>>();
            // `kind` belongs here too — this is the ONLY path with a
            // `type_env`, so it is the only one that can answer for a
            // VARIABLE. Without it `kind(x)` fell through to the shared
            // compiler's hard-coded 8.
            if matches!(
                lowered.as_str(),
                "kind"
                    | "bit_size"
                    | "storage_size"
                    | "precision"
                    | "range"
                    | "digits"
                    | "maxexponent"
                    | "minexponent"
            ) {
                let positional_args = args
                    .iter()
                    .filter(|arg| {
                        arg.name
                            .as_deref()
                            .is_none_or(|name| !name.eq_ignore_ascii_case("kind"))
                    })
                    .map(|arg| arg.value.clone())
                    .collect::<Vec<_>>();
                if let Some(first) = positional_args.first() {
                    if let Some(folded) = fold_fortran_type_inquiry(&lowered, first, type_env, None)
                    {
                        *expr = folded;
                    }
                }
            }
            match lowered.as_str() {
                // F2003 type inquiry. Both ask about the DYNAMIC type of the
                // first argument, and for a NON-POLYMORPHIC entity that is its
                // declared type — which is exactly what this pass carries and
                // nothing else does.
                //
                // `extends_type_of(a, mold)` is "a's type IS mold's type or an
                // extension of it", which `IsType` without the exact-match
                // prefix already means, so the hierarchy answer comes from the
                // machinery `class is` uses. `same_type_as` demands the exact
                // type, and the exact-match prefix is only ever resolved inside
                // a marked `select type` chain — so it is answered here from
                // the two declared names, which is the whole truth when neither
                // entity is polymorphic.
                "extends_type_of" | "same_type_as" if positional.len() == 2 => {
                    let arg_hint = fortran_type_hint_for_expr(&positional[0], type_env);
                    let mold_hint = fortran_type_hint_for_expr(&positional[1], type_env);
                    if let Some(mold) = mold_hint {
                        let mold_name = fortran_canonical_select_type_name(&mold);
                        if lowered == "extends_type_of" {
                            *expr = Expression::new(ExprKind::IsType {
                                expr: Box::new(positional[0].clone()),
                                type_name: mold_name,
                            });
                        } else if let Some(arg) = arg_hint {
                            // ⛔ Only when NEITHER is polymorphic: a `class(T)`
                            // entity's declared type is an upper bound, not an
                            // answer, and folding it would state a fact the
                            // program has not established.
                            let polymorphic = arg.trim().to_ascii_lowercase().starts_with("class(")
                                || mold.trim().to_ascii_lowercase().starts_with("class(");
                            if !polymorphic {
                                let same =
                                    fortran_canonical_select_type_name(&arg) == mold_name;
                                *expr = Expression::new(ExprKind::Lit(Literal::Bool(same)));
                            }
                        }
                    }
                }
                "nearest" if positional.len() == 2 => {
                    let wide = fortran_expr_is_kind8_real(&positional[0], type_env);
                    let up = if wide {
                        "__vybe_next_up64"
                    } else {
                        "__vybe_next_up32"
                    };
                    let down = if wide {
                        "__vybe_next_dn64"
                    } else {
                        "__vybe_next_dn32"
                    };
                    *expr = fortran_ternary(
                        fortran_bin(BinOp::Lt, positional[1].clone(), Expression::float(0.0)),
                        fortran_call(down, vec![positional[0].clone()]),
                        fortran_call(up, vec![positional[0].clone()]),
                    );
                }
                // The bare `spacing` builtin row IS the kind-4 lane; only a
                // kind-8 operand needs redirecting.
                "spacing"
                    if positional.len() == 1
                        && fortran_expr_is_kind8_real(&positional[0], type_env) =>
                {
                    *expr = fortran_call("__vybe_ulp64", vec![positional[0].clone()]);
                }
                _ => {}
            }
        }
        ExprKind::Binary { left, right, .. } => {
            lower_fortran_type_inquiry_in_expr(left, type_env);
            lower_fortran_type_inquiry_in_expr(right, type_env);
        }
        ExprKind::Unary { expr: inner, .. } => {
            lower_fortran_type_inquiry_in_expr(inner, type_env);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            lower_fortran_type_inquiry_in_expr(cond, type_env);
            lower_fortran_type_inquiry_in_expr(then, type_env);
            lower_fortran_type_inquiry_in_expr(else_, type_env);
        }
        ExprKind::Member { object, .. } => {
            lower_fortran_type_inquiry_in_expr(object, type_env);
        }
        ExprKind::Index { object, index, .. } => {
            lower_fortran_type_inquiry_in_expr(object, type_env);
            lower_fortran_type_inquiry_in_expr(index, type_env);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                lower_fortran_type_inquiry_in_expr(&mut element.value, type_env);
            }
        }
        // A `print *, a, b` with MORE THAN ONE item lowers to an
        // interpolation, so every type inquiry inside a multi-item print was
        // left unfolded and reached the runtime as a call to a function that
        // does not exist. A single-item print takes a different path, which is
        // why `print *, storage_size(s)` worked and adding a second item broke
        // both of them.
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(inner) | InterpolPart::Formatted(inner, _) = part {
                    lower_fortran_type_inquiry_in_expr(inner, type_env);
                }
            }
        }
        _ => {}
    }
}

fn lower_fortran_type_inquiry_in_statement(
    statement: &mut Statement,
    type_env: &HashMap<String, String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            lower_fortran_type_inquiry_in_expr(expr, type_env);
        }
        StmtKind::Return(expr) => {
            if let Some(value) = expr {
                lower_fortran_type_inquiry_in_expr(value, type_env);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(value) = expr {
                lower_fortran_type_inquiry_in_expr(value, type_env);
            }
            if let Some(value) = cause {
                lower_fortran_type_inquiry_in_expr(value, type_env);
            }
        }
        StmtKind::PrintFile { items, .. } | StmtKind::WriteFile { items, .. } => {
            for item in items {
                lower_fortran_type_inquiry_in_expr(item, type_env);
            }
        }
        StmtKind::Assign { value, .. } => {
            lower_fortran_type_inquiry_in_expr(value, type_env);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            lower_fortran_type_inquiry_in_expr(cond, type_env);
            for stmt in then_body.iter_mut() {
                lower_fortran_type_inquiry_in_statement(stmt, type_env);
            }
            for (branch_cond, branch_body) in elifs.iter_mut() {
                lower_fortran_type_inquiry_in_expr(branch_cond, type_env);
                for stmt in branch_body.iter_mut() {
                    lower_fortran_type_inquiry_in_statement(stmt, type_env);
                }
            }
            if let Some(body) = else_body {
                for stmt in body.iter_mut() {
                    lower_fortran_type_inquiry_in_statement(stmt, type_env);
                }
            }
        }
        _ => {}
    }
}

fn build_fortran_sign_expr(magnitude: Expression, sign_source: Expression) -> Expression {
    let abs_magnitude = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("abs")),
        args: vec![Argument::positional(magnitude)],
        optional: false,
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(sign_source),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(abs_magnitude.clone()),
        else_: Box::new(Expression::new(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: Box::new(abs_magnitude),
        })),
    })
}

fn build_fortran_modulo_expr(value: Expression, modulus: Expression) -> Expression {
    let quotient = Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(value.clone()),
        right: Box::new(modulus.clone()),
    });
    let floored = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("floor")),
        args: vec![Argument::positional(quotient)],
        optional: false,
    });
    Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(value),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(modulus),
            right: Box::new(floored),
        })),
    })
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
        return Ok(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(walk_expr(inner.remove(1))?),
        }));
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
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }));
        }
        // `+` is a no-op; the `primary` form is the bare value.
        return walk_expr(inner.remove(0));
    }
    // `power = { concat ~ ("**" ~ concat)? }` — inline `**` literal,
    // so inner has either [base] or [base, exponent]. Apply Pow when 2.
    if rule == Rule::power && inner.len() == 2 {
        let base = walk_expr(inner.remove(0))?;
        let exp = walk_expr(inner.remove(0))?;
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
    if inner.len() == 1 {
        return walk_expr(inner.remove(0));
    }
    if inner.len() >= 3 {
        let mut result = walk_expr(inner.remove(0))?;
        let mut i = 0;
        while i + 1 < inner.len() {
            let op_text = inner[i].as_str().to_lowercase();
            let op = to_binop(&inner[i]);
            let right = walk_expr(inner[i + 1].clone())?;
            result = Expression::new(ExprKind::Binary {
                left: Box::new(result),
                op,
                right: Box::new(right),
            });
            if op_text.trim_start().starts_with(".neqv.") {
                result = Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(result),
                });
            }
            i += 2;
        }
        return Ok(result);
    }
    if inner.is_empty() {
        return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
    }
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
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "**" => BinOp::Pow,
        "//" => BinOp::Concat,
        "==" | ".eq." => BinOp::Eq,
        "/=" | ".ne." => BinOp::NotEq,
        "<" | ".lt." => BinOp::Lt,
        ">" | ".gt." => BinOp::Gt,
        "<=" | ".le." => BinOp::LtEq,
        ">=" | ".ge." => BinOp::GtEq,
        ".and." => BinOp::And,
        ".or." => BinOp::Or,
        ".eqv." => BinOp::Eqv,
        ".neqv." => BinOp::Eqv,
        _ => BinOp::Add,
    }
}

fn meaningful(pair: &Pair<Rule>) -> bool {
    !matches!(pair.as_rule(), Rule::NEWLINE | Rule::EOI)
}

fn is_expr_rule(r: Rule) -> bool {
    matches!(
        r,
        Rule::expression
            | Rule::logical_or
            | Rule::logical_and
            | Rule::comparison
            | Rule::addition
            | Rule::multiplication
            | Rule::power
            | Rule::concat
            | Rule::unary
            | Rule::primary_expr
            | Rule::literal
            | Rule::number_literal
            | Rule::string_literal
            | Rule::identifier
            | Rule::function_call_or_subscript
            | Rule::logical_literal
            | Rule::logical_not
    )
}

fn to_class_member(stmt: Statement) -> ClassMember {
    match stmt.kind {
        StmtKind::ClassDecl { .. }
        | StmtKind::StructDecl { .. }
        | StmtKind::EnumDecl { .. }
        | StmtKind::InterfaceDecl { .. }
        | StmtKind::ModuleDecl { .. } => ClassMember::NestedType(Box::new(stmt)),
        StmtKind::FunctionDecl { .. } => ClassMember::Method(Box::new(stmt)),
        StmtKind::VarDecl {
            ref declarations, ..
        } => {
            if let Some(d) = declarations.first() {
                if let BindingPattern::Ident(name) = &d.pattern {
                    let field_type_hint = d.type_hint.as_ref().map(|type_hint| {
                        fortran_array_type_hint(type_hint, d.array_bounds.as_deref())
                    });
                    return ClassMember::Field {
                        name: name.clone(),
                        type_hint: field_type_hint,
                        init: d.init.clone(),
                        modifiers: Modifiers::default(),
                        with_events: false,
                        array_bounds: d.array_bounds.clone(),
                        storage: None,
                    };
                }
            }
            ClassMember::Method(Box::new(stmt))
        }
        _ => ClassMember::Method(Box::new(stmt)),
    }
}
