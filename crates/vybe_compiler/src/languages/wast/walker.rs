// ============================================================================
// WAST / WAT Walker — pest parse tree → common AST
// ============================================================================
// WAT (WebAssembly Text Format) is the human-readable form of WASM binary.
// WAST is a superset that adds script commands: assert_return, assert_trap,
// assert_invalid, invoke, register, etc.
//
// Mapping strategy:
//   (module id? field*) → ClassDecl (static methods = funcs, fields = globals)
//   (func id? typeuse local* instr*) → FunctionDecl (static)
//   WAT instruction → Call(instr_name_with_dots_replaced, args)
//     with common ops mapped to BinOp/UnaryOp/Ternary where possible
//   WAST script commands → Call(__wast_assert_return, __wast_assert_trap, …)
// ============================================================================

use pest::Parser;
use pest::iterators::Pair;
use super::{WastParser, Rule};
use crate::ast::*;

// ── Entry point ──────────────────────────────────────────────────────────────

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = WastParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();

    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for cmd in top.into_inner() {
                    match cmd.as_rule() {
                        Rule::EOI => {}
                        _ => walk_script_cmd(cmd, &mut body)?,
                    }
                }
            }
            Rule::EOI => {}
            _ => walk_script_cmd(top, &mut body)?,
        }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
    })
}

// ── Script commands ───────────────────────────────────────────────────────────

fn walk_script_cmd(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::script_cmd => {
            let inner = pair.into_inner().next().ok_or("Empty script_cmd")?;
            walk_script_cmd(inner, body)
        }
        Rule::module                    => { body.extend(walk_module(pair)?); Ok(()) }
        Rule::assert_return             => { body.push(walk_assert_return(pair)?);    Ok(()) }
        Rule::assert_trap               => { body.push(walk_assert_trap(pair)?);      Ok(()) }
        Rule::assert_instantiation_trap => { body.push(walk_assert_trap(pair)?);      Ok(()) }
        Rule::assert_invalid            => { body.push(walk_assert_generic(pair, "__wast_assert_invalid")?);    Ok(()) }
        Rule::assert_malformed          => { body.push(walk_assert_generic(pair, "__wast_assert_malformed")?);  Ok(()) }
        Rule::assert_unlinkable         => { body.push(walk_assert_generic(pair, "__wast_assert_unlinkable")?); Ok(()) }
        Rule::assert_exhaustion         => { body.push(walk_assert_generic(pair, "__wast_assert_exhaustion")?); Ok(()) }
        Rule::assert_suspension         => { body.push(walk_assert_generic(pair, "__wast_assert_suspension")?); Ok(()) }
        Rule::invoke_cmd                => { body.push(walk_invoke_cmd(pair)?);       Ok(()) }
        Rule::register_cmd              => { body.push(walk_register_cmd(pair)?);     Ok(()) }
        Rule::get_cmd                   => { body.push(walk_get_cmd(pair)?);          Ok(()) }
        _ => Ok(()),
    }
}

// ── Module ────────────────────────────────────────────────────────────────────

fn walk_module(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut module_name: Option<String> = None;
    let mut members: Vec<ClassMember> = Vec::new();
    let mut top_stmts: Vec<Statement> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {
                module_name = Some(child.as_str()[1..].to_string());
            }
            Rule::module_field => {
                let inner = child.into_inner().next().ok_or("Empty module_field")?;
                match inner.as_rule() {
                    Rule::func_field => {
                        members.push(ClassMember::Method(Box::new(walk_func_field(inner)?)));
                    }
                    Rule::import_field => {
                        top_stmts.push(walk_import_field(inner)?);
                    }
                    Rule::export_field => {
                        let expr = walk_export_field(inner)?;
                        top_stmts.push(Statement::new(StmtKind::Expr(expr)));
                    }
                    Rule::global_field => {
                        let (name, init) = walk_global_field(inner)?;
                        members.push(ClassMember::Field {
                            name,
                            type_hint: None,
                            init: Some(init),
                            modifiers: Modifiers { is_static: true, ..Default::default() },
                            with_events: false,
                            array_bounds: None,
                        });
                    }
                    Rule::start_field => {
                        top_stmts.push(Statement::new(StmtKind::Expr(walk_start_field(inner)?)));
                    }
                    _ => {} // table, memory, elem, data, type — structural metadata
                }
            }
            _ => {}
        }
    }

    let name = module_name.unwrap_or_else(|| "__wasm_module".to_string());
    let class = Statement::with_span(
        StmtKind::ClassDecl {
            name,
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span,
    );

    let mut result = vec![class];
    result.extend(top_stmts);
    Ok(result)
}

// ── Function field ────────────────────────────────────────────────────────────

fn walk_func_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut export_names: Vec<String> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {
                func_name = child.as_str()[1..].to_string();
            }
            Rule::export_inline => {
                if let Some(s) = child.into_inner().find(|p| p.as_rule() == Rule::string) {
                    export_names.push(unquote(s.as_str()));
                }
            }
            Rule::import_inline => {} // stub — body stays empty
            Rule::typeuse => {
                params = walk_typeuse_params(child)?;
            }
            Rule::local => {
                body.extend(walk_local(child)?);
            }
            Rule::instr => {
                body.push(walk_instr_as_stmt(child)?);
            }
            _ => {}
        }
    }

    if func_name.is_empty() {
        func_name = export_names.first()
            .cloned()
            .unwrap_or_else(|| "__wasm_func".to_string());
    }

    // WAT functions implicitly return the last value left on the stack
    apply_implicit_return(&mut body);

    let mut modifiers = Modifiers::default();
    modifiers.is_static = true;

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name: func_name,
            params,
            return_type: None,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span,
    ))
}

fn walk_typeuse_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::param {
            params.extend(walk_param(child)?);
        }
    }
    Ok(params)
}

fn walk_param(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id       => name = Some(child.as_str()[1..].to_string()),
            Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }

    if types.is_empty() { return Ok(Vec::new()); }

    if let Some(n) = name {
        return Ok(vec![Param {
            name: n,
            type_hint: types.into_iter().next(),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false,
        }]);
    }

    Ok(types.into_iter().enumerate().map(|(i, t)| Param {
        name: format!("p{}", i),
        type_hint: Some(t),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false,
    }).collect())
}

fn walk_local(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id       => name = Some(child.as_str()[1..].to_string()),
            Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }

    Ok(types.iter().enumerate().map(|(i, t)| {
        let var_name = name.clone().unwrap_or_else(|| format!("local{}", i));
        let init = match t.as_str() {
            "i32" | "i64" => Expression::int(0),
            "f32" | "f64" => Expression::float(0.0),
            _              => Expression::null(),
        };
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(var_name),
                type_hint: Some(t.clone()),
                init: Some(init),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        })
    }).collect())
}

// ── Instructions ──────────────────────────────────────────────────────────────

fn walk_instr_as_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let expr = walk_instr_as_expr(pair)?;
    Ok(Statement::with_span(StmtKind::Expr(expr), span))
}

fn walk_instr_as_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr(inner, span),
        Rule::plain_instr  => walk_plain_instr(inner, span),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
    }
}

fn walk_folded_instr(pair: Pair<Rule>, span: Span) -> Result<Expression, String> {
    let mut name = String::new();
    let mut args: Vec<Expression> = Vec::new();
    let mut then_exprs: Vec<Expression> = Vec::new();
    let mut else_exprs: Vec<Expression> = Vec::new();
    let mut has_then = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::instr_name  => name = child.as_str().to_string(),
            Rule::instr_arg   => args.push(walk_instr_arg(child)?),
            Rule::instr       => args.push(walk_instr_as_expr(child)?),
            Rule::then_block  => {
                has_then = true;
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        then_exprs.push(walk_instr_as_expr(sub)?);
                    }
                }
            }
            Rule::else_block  => {
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        else_exprs.push(walk_instr_as_expr(sub)?);
                    }
                }
            }
            // try/catch — catch bodies become extra args to __wasm_try
            Rule::catch_block | Rule::catch_all_block => {
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        args.push(walk_instr_as_expr(sub)?);
                    }
                }
            }
            _ => {}
        }
    }

    // (try … (catch …)) → __wasm_try(…)
    if name == "try" {
        return Ok(make_call("__wasm_try", args, span));
    }

    // (if cond (then ...) (else ...)) → ternary
    if name == "if" || has_then {
        let cond = args.into_iter().next().unwrap_or(Expression::bool(false));
        let then_val = then_exprs.into_iter().last().unwrap_or(Expression::null());
        let else_val = else_exprs.into_iter().last().unwrap_or(Expression::null());
        return Ok(Expression::with_span(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        }, span));
    }

    map_instr_to_ast(name, args, span)
}

fn walk_plain_instr(pair: Pair<Rule>, span: Span) -> Result<Expression, String> {
    let mut name = String::new();
    let mut args: Vec<Expression> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::instr_name => name = child.as_str().to_string(),
            Rule::instr_arg  => args.push(walk_instr_arg(child)?),
            _ => {}
        }
    }
    map_instr_to_ast(name, args, span)
}

/// Map a WAT instruction + args to the most appropriate common AST node.
fn map_instr_to_ast(name: String, mut args: Vec<Expression>, span: Span) -> Result<Expression, String> {
    match name.as_str() {
        // ── Constants → literals ──────────────────────────────────────────
        "i32.const" | "i64.const" => {
            Ok(args.into_iter().next().unwrap_or(Expression::int(0)))
        }
        "f32.const" | "f64.const" => {
            Ok(args.into_iter().next().unwrap_or(Expression::float(0.0)))
        }

        // ── Local/global get → Ident ──────────────────────────────────────
        "local.get" | "global.get" => {
            let idx = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(match &idx.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => Expression::with_span(ExprKind::Ident(format!("p{}", i)), span),
                _ => idx,
            })
        }

        // ── Local/global set → Assign ─────────────────────────────────────
        "local.set" | "global.set" => {
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target = match &target_raw.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => Expression::with_span(ExprKind::Ident(format!("p{}", i)), span),
                _ => target_raw,
            };
            Ok(Expression::with_span(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            }, span))
        }

        // ── local.tee → assign + return value ────────────────────────────
        "local.tee" => {
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target_name = match &target_raw.kind {
                ExprKind::Ident(n) => n.clone(),
                ExprKind::Lit(Literal::Int(i)) => format!("p{}", i),
                _ => "__tee_tmp".to_string(),
            };
            let assign = Expression::with_span(ExprKind::Assign {
                target: Box::new(Expression::ident(&target_name)),
                value: Box::new(value),
            }, span);
            Ok(Expression::with_span(ExprKind::Sequence(vec![
                assign,
                Expression::ident(&target_name),
            ]), span))
        }

        // ── Binary arithmetic → BinOp ─────────────────────────────────────
        "i32.add" | "i64.add" | "f32.add" | "f64.add" => bin_op(args, BinOp::Add, span),
        "i32.sub" | "i64.sub" | "f32.sub" | "f64.sub" => bin_op(args, BinOp::Sub, span),
        "i32.mul" | "i64.mul" | "f32.mul" | "f64.mul" => bin_op(args, BinOp::Mul, span),
        "i32.div_s" | "i32.div_u" | "i64.div_s" | "i64.div_u"
        | "f32.div" | "f64.div"                        => bin_op(args, BinOp::Div, span),
        "i32.rem_s" | "i32.rem_u" | "i64.rem_s" | "i64.rem_u" => bin_op(args, BinOp::Mod, span),
        "i32.and"   | "i64.and"                        => bin_op(args, BinOp::BitAnd, span),
        "i32.or"    | "i64.or"                         => bin_op(args, BinOp::BitOr, span),
        "i32.xor"   | "i64.xor"                        => bin_op(args, BinOp::BitXor, span),
        "i32.shl"   | "i64.shl"                        => bin_op(args, BinOp::Shl, span),
        "i32.shr_s" | "i32.shr_u" | "i64.shr_s" | "i64.shr_u" => bin_op(args, BinOp::Shr, span),

        // ── Comparisons → BinOp ───────────────────────────────────────────
        "i32.eq" | "i64.eq" | "f32.eq" | "f64.eq"     => bin_op(args, BinOp::Eq, span),
        "i32.ne" | "i64.ne" | "f32.ne" | "f64.ne"     => bin_op(args, BinOp::NotEq, span),
        "i32.lt_s" | "i32.lt_u" | "i64.lt_s" | "i64.lt_u" | "f32.lt" | "f64.lt" => bin_op(args, BinOp::Lt, span),
        "i32.gt_s" | "i32.gt_u" | "i64.gt_s" | "i64.gt_u" | "f32.gt" | "f64.gt" => bin_op(args, BinOp::Gt, span),
        "i32.le_s" | "i32.le_u" | "i64.le_s" | "i64.le_u" | "f32.le" | "f64.le" => bin_op(args, BinOp::LtEq, span),
        "i32.ge_s" | "i32.ge_u" | "i64.ge_s" | "i64.ge_u" | "f32.ge" | "f64.ge" => bin_op(args, BinOp::GtEq, span),

        // ── eqz → == 0 ───────────────────────────────────────────────────
        "i32.eqz" | "i64.eqz" => {
            let operand = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(Expression::with_span(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(operand),
                right: Box::new(Expression::int(0)),
            }, span))
        }

        // ── Unary ─────────────────────────────────────────────────────────
        "f32.neg" | "f64.neg" => {
            let operand = args.into_iter().next().unwrap_or(Expression::float(0.0));
            Ok(Expression::with_span(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }, span))
        }

        // ── select → ternary ──────────────────────────────────────────────
        // WAT: select val1 val2 cond → cond != 0 ? val1 : val2
        "select" => {
            let mut it = args.into_iter();
            let val1 = it.next().unwrap_or(Expression::null());
            let val2 = it.next().unwrap_or(Expression::null());
            let cond = it.next().unwrap_or(Expression::bool(false));
            Ok(Expression::with_span(ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(val1),
                else_: Box::new(val2),
            }, span))
        }

        // ── drop → evaluate and discard (identity) ────────────────────────
        "drop" => Ok(args.into_iter().next().unwrap_or(Expression::null())),

        // ── nop → null literal ────────────────────────────────────────────
        "nop" => Ok(Expression::with_span(ExprKind::Lit(Literal::Null), span)),

        // ── unreachable → throw ───────────────────────────────────────────
        // Emitted as a call to __wasm_unreachable() so the compiler can map
        // it to a throw without needing ExprKind::Throw (which doesn't exist).
        "unreachable" => Ok(make_call("__wasm_unreachable", vec![], span)),

        // ── call → Call(callee, args) ─────────────────────────────────────
        "call" => {
            let mut it = args.into_iter();
            let callee = it.next().unwrap_or(Expression::null());
            let call_args: Vec<Expression> = it.collect();
            Ok(Expression::with_span(ExprKind::Call {
                callee: Box::new(callee),
                args: call_args.into_iter().map(Argument::positional).collect(),
                optional: false,
            }, span))
        }

        // ── return → __wasm_return(val?) ──────────────────────────────────
        "return" => Ok(make_call("__wasm_return", args, span)),

        // ── br / br_if / br_table → calls ────────────────────────────────
        "br"       => Ok(make_call("__wasm_br",       args, span)),
        "br_if"    => Ok(make_call("__wasm_br_if",    args, span)),
        "br_table" => Ok(make_call("__wasm_br_table", args, span)),

        // ── Everything else → generic call with dots→underscores ──────────
        _ => Ok(make_call(&name.replace('.', "_"), args, span)),
    }
}

// ── Instruction argument ──────────────────────────────────────────────────────

fn walk_instr_arg(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty instr_arg")?;
    match inner.as_rule() {
        Rule::float         => Ok(parse_float(inner.as_str())),
        Rule::integer       => Ok(parse_integer(inner.as_str())),
        Rule::string        => Ok(Expression::string(&unquote(inner.as_str()))),
        Rule::id            => Ok(Expression::ident(&inner.as_str()[1..])),
        Rule::val_type      => Ok(Expression::string(inner.as_str())),
        Rule::bare_val_type => Ok(Expression::string(inner.as_str())),
        Rule::bare_lane_type=> Ok(Expression::string(inner.as_str())),
        Rule::mem_arg       => Ok(Expression::string(inner.as_str())),
        Rule::folded_instr  => walk_folded_instr(inner, Span::default()),
        Rule::val_lane_type => Ok(Expression::string(inner.as_str())),
        _                   => Ok(Expression::null()),
    }
}

fn walk_index(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty index")?;
    match inner.as_rule() {
        Rule::integer => Ok(parse_integer(inner.as_str())),
        Rule::id      => Ok(Expression::ident(&inner.as_str()[1..])),
        _             => Ok(Expression::null()),
    }
}

// ── Module fields ─────────────────────────────────────────────────────────────

fn walk_import_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut module_str = String::new();
    let mut name_str = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            let s = unquote(child.as_str());
            if module_str.is_empty() { module_str = s; }
            else if name_str.is_empty() { name_str = s; }
        }
    }
    Ok(Statement::new(StmtKind::Expr(make_call(
        "__wasm_import",
        vec![Expression::string(&module_str), Expression::string(&name_str)],
        Span::default(),
    ))))
}

fn walk_export_field(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut export_name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            export_name = unquote(child.as_str());
            break;
        }
    }
    Ok(make_call("__wasm_export", vec![Expression::string(&export_name)], Span::default()))
}

fn walk_global_field(pair: Pair<Rule>) -> Result<(String, Expression), String> {
    let mut name = "__wasm_global".to_string();
    let mut init = Expression::int(0);
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id    => name = child.as_str()[1..].to_string(),
            Rule::instr => init = walk_instr_as_expr(child)?,
            _ => {}
        }
    }
    Ok((name, init))
}

fn walk_start_field(pair: Pair<Rule>) -> Result<Expression, String> {
    let idx = pair.into_inner().next().ok_or("Empty start field")?;
    let callee = walk_index(idx)?;
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: Vec::new(),
        optional: false,
    }))
}

// ── WAST script commands ──────────────────────────────────────────────────────

/// (invoke id? "func" expr*) → Call(func, args)
fn walk_invoke_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut args: Vec<Expression> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id     => {} // optional module id — ignore
            Rule::string => {
                if func_name.is_empty() { func_name = unquote(child.as_str()); }
            }
            Rule::expr   => args.push(walk_const_expr(child)?),
            _ => {}
        }
    }

    Ok(Statement::with_span(StmtKind::Expr(Expression::with_span(ExprKind::Call {
        callee: Box::new(Expression::ident(&func_name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    }, span)), span))
}

/// (assert_return action result*) → __wast_assert_return(action_result, expected...)
fn walk_assert_return(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    let mut expected: Vec<Expression> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action     => action_expr = Some(walk_action(child)?),
            Rule::result_val => expected.push(walk_const_expr(child)?),
            _ => {}
        }
    }

    let mut args = Vec::new();
    if let Some(a) = action_expr { args.push(a); }
    args.extend(expected);

    Ok(Statement::with_span(StmtKind::Expr(
        make_call("__wast_assert_return", args, span)
    ), span))
}

/// (assert_trap action "message") → __wast_assert_trap(action_result, "message")
fn walk_assert_trap(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    let mut message = String::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action => action_expr = Some(walk_action(child)?),
            Rule::string => message = unquote(child.as_str()),
            _ => {}
        }
    }

    let mut args = Vec::new();
    if let Some(a) = action_expr { args.push(a); }
    args.push(Expression::string(&message));

    Ok(Statement::with_span(StmtKind::Expr(
        make_call("__wast_assert_trap", args, span)
    ), span))
}

fn walk_assert_generic(pair: Pair<Rule>, fn_name: &str) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut message = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            message = unquote(child.as_str());
        }
    }
    Ok(Statement::with_span(StmtKind::Expr(
        make_call(fn_name, vec![Expression::string(&message)], span)
    ), span))
}

fn walk_register_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string { name = unquote(child.as_str()); break; }
    }
    Ok(Statement::with_span(StmtKind::Expr(
        make_call("__wasm_register", vec![Expression::string(&name)], span)
    ), span))
}

fn walk_get_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string { name = unquote(child.as_str()); break; }
    }
    Ok(Statement::with_span(StmtKind::Expr(
        make_call("__wasm_get", vec![Expression::string(&name)], span)
    ), span))
}

fn walk_action(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty action")?;
    match inner.as_rule() {
        Rule::invoke_cmd => {
            let stmt = walk_invoke_cmd(inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()),
            }
        }
        Rule::get_cmd => {
            let stmt = walk_get_cmd(inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()),
            }
        }
        _ => Ok(Expression::null()),
    }
}

fn walk_const_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::integer       => return Ok(parse_integer(child.as_str())),
            Rule::float         => return Ok(parse_float(child.as_str())),
            Rule::string        => return Ok(Expression::string(&unquote(child.as_str()))),
            Rule::val_lane_type => {} // v128 lane type — skip, treat as null
            _ => {}
        }
    }
    Ok(Expression::null())
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn make_call(name: &str, args: Vec<Expression>, span: Span) -> Expression {
    Expression::with_span(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    }, span)
}

fn bin_op(mut args: Vec<Expression>, op: BinOp, span: Span) -> Result<Expression, String> {
    // WAT stack order: operands are pushed left-to-right, so args[0]=left, args[1]=right
    let right = if args.len() >= 2 { args.remove(1) } else { Expression::int(0) };
    let left  = args.into_iter().next().unwrap_or(Expression::int(0));
    Ok(Expression::with_span(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    }, span))
}

/// If the last statement in a function body is an Expr, wrap it in Return.
/// WAT functions implicitly return the last value left on the stack.
fn apply_implicit_return(body: &mut Vec<Statement>) {
    if body.is_empty() { return; }
    if let Some(last) = body.last_mut() {
        if let StmtKind::Expr(ref e) = last.kind.clone() {
            // Don't double-wrap __wasm_return calls
            if let ExprKind::Call { ref callee, .. } = e.kind {
                if let ExprKind::Ident(ref n) = callee.kind {
                    if n == "__wasm_return" { return; }
                }
            }
            last.kind = StmtKind::Return(Some(e.clone()));
        }
    }
}

fn parse_integer(s: &str) -> Expression {
    let s = s.trim().replace('_', "");
    // Handle signed hex: -0x..., +0x..., 0x...
    let (neg, digits) = if s.starts_with("-0x") || s.starts_with("-0X") {
        (true, &s[3..])
    } else if s.starts_with("0x") || s.starts_with("0X") {
        (false, &s[2..])
    } else if s.starts_with("+0x") || s.starts_with("+0X") {
        (false, &s[3..])
    } else {
        return Expression::int(s.parse::<i64>().unwrap_or(0));
    };
    let v = i64::from_str_radix(digits, 16).unwrap_or(0);
    Expression::int(if neg { -v } else { v })
}

fn parse_float(s: &str) -> Expression {
    let s = s.trim();
    match s {
        "inf" | "+inf"               => Expression::float(f64::INFINITY),
        "-inf"                       => Expression::float(f64::NEG_INFINITY),
        "nan" | "+nan" | "-nan"      => Expression::float(f64::NAN),
        _ if s.contains("nan:0x")    => Expression::float(f64::NAN),
        _                            => Expression::float(s.parse::<f64>().unwrap_or(0.0)),
    }
}

fn unquote(s: &str) -> String {
    let s = s.trim();
    if s.len() >= 2 && s.starts_with('"') && s.ends_with('"') {
        s[1..s.len()-1]
            .replace("\\n", "\n")
            .replace("\\t", "\t")
            .replace("\\r", "\r")
            .replace("\\\\", "\\")
            .replace("\\\"", "\"")
    } else {
        s.to_string()
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let start = pair.as_span().start_pos().line_col();
    let end   = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32,
        start_col:  start.1 as u32,
        end_line:   end.0 as u32,
        end_col:    end.1 as u32,
    }
}
