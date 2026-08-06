// ============================================================================
// WAST / WAT Walker — pest parse tree → common AST
// ============================================================================
// WAT (WebAssembly Text Format) is the human-readable form of WASM binary.
// WAST is a superset that adds script commands: assert_return, assert_trap,
// assert_invalid, invoke, register, etc.
//
// Mapping strategy:
//   (module id? field*) → ClassDecl (static methods = funcs, globals = pre_stmts)
//   (func id? typeuse local* instr*) → FunctionDecl (static)
//   WAT instruction → common AST:
//     block $l (...)  → StmtKind::Labeled { label: l, Block([...]) }
//     loop  $l (...)  → StmtKind::Labeled { label: l, While(true, [...]) }
//     br $l           → StmtKind::Break(Label(l))  if l is a block label
//                     → StmtKind::Continue(Label(l)) if l is a loop label
//     br_if $l cond   → If(cond, [Break/Continue(l)])
//     return val?     → StmtKind::Return(val)
//     unreachable     → StmtKind::Throw (WASM trap)
//     if (then)(else) → ExprKind::Ternary
//     binary ops      → ExprKind::Binary
//     call $f args    → ExprKind::Call
//     local.get $x    → ExprKind::Ident
//     local.set $x v  → ExprKind::Assign
//     i32.const N     → ExprKind::Lit
//     everything else → Call(name_with_underscores, args)
//   WAST script cmds  → Call(__wast_assert_return / __wast_assert_trap / …)
// ============================================================================

use super::{Rule, WastParser};
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;
use std::collections::HashMap;
use vybe_ast::*;

thread_local! {
    static FUNC_INDEX_ARITIES: RefCell<Vec<usize>> = RefCell::new(Vec::new());
    static FUNC_NAME_ARITIES: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // How many results each function yields, by index and by name. A `call` to a
    // 0-result function is a void statement, not a value pushed on the stack; if
    // it were pushed it would be deferred to the block's stack flush and run out
    // of order (WASM linear code runs calls in place).
    static FUNC_INDEX_RESULTS: RefCell<Vec<usize>> = RefCell::new(Vec::new());
    static FUNC_NAME_RESULTS: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // The result count of the function currently being folded, so a `return`
    // inside a multi-value function (result_count >= 2) pops N values and
    // reraises them as a uniform tuple `return` (the multi-value-tuple ABI).
    static CURRENT_FN_RESULTS: RefCell<usize> = const { RefCell::new(0) };
    // type name → number of fields (for struct.new arity)
    static STRUCT_FIELD_COUNTS: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // GC struct types in declaration order: (name, parent name, field count).
    // Registered into the type table so `ref.test`/`ref.cast`/`br_on_cast`
    // resolve identity + subtyping (the parent link IS the subtype edge).
    static STRUCT_TYPES: RefCell<Vec<(String, Option<String>, usize)>> = RefCell::new(Vec::new());
    // struct type name → field storage types in order (`i8`/`i16`/`i32`/`i64`/
    // `f32`/`f64`/ref…). Drives `struct.new_default` field defaults and
    // `struct.get_s`/`get_u` packed sign/zero-extension.
    static STRUCT_FIELD_TYPES: RefCell<HashMap<String, Vec<String>>> = RefCell::new(HashMap::new());
    // func-type name → number of params (for call_ref/call_indirect arity)
    static TYPE_FUNC_PARAMS: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // func-type name → number of results (the other half of the type shape the
    // `call_indirect` runtime check compares against the callee).
    static TYPE_FUNC_RESULTS: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // table name (`$t1`) → table index, in module declaration order. Lets `elem`
    // segments populate, and `call_indirect $t` dispatch through, a NAMED table.
    static TABLE_NAME_INDEX: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // memory name (`$m2`) → memory index, in module declaration order. Lets
    // multi-memory `i32.load/store $m`, `memory.size $m`, and `(data (memory $m))`
    // target the right linear memory.
    static MEMORY_NAME_INDEX: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // Global declaration-order index → the binding name it lowers to (its `$id`,
    // or a synthetic `__wasm_global_<i>` when unnamed). Lets `global.get N` /
    // `global.set N` by NUMERIC index resolve to the right binding (globals do
    // not share the local/param `p<i>` name space).
    static GLOBAL_INDEX_NAME: RefCell<Vec<String>> = RefCell::new(Vec::new());
    // array type name → element storage type (`i8`/`i16` packed, or `i32`/…).
    // `array.get_s`/`get_u` on a packed array sign-extend/zero-extend by width.
    static ARRAY_ELEM_TYPE: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    // Running element-segment index (declaration order), so a passive `(elem …)`
    // registers under the same index `table.init`/`array.new_elem` reference.
    static ELEM_SEG_COUNTER: RefCell<usize> = const { RefCell::new(0) };
    // Module functions compile to static methods of this class; a `call $f` to a
    // defined function is reached as `ClassName.f(...)`.
    static MODULE_CLASS_NAME: RefCell<String> = RefCell::new(String::new());
    // Names of functions DEFINED in the module (not imports) — call targets that
    // must be qualified with the module class name. Imports resolve via the
    // profile builtin table by their local id, so they are excluded.
    static DEFINED_FUNC_NAMES: RefCell<std::collections::HashSet<String>> =
        RefCell::new(std::collections::HashSet::new());
    // Exported function name → the static-method name it maps to on the module
    // class, so a WAST script `(invoke "name" …)` resolves to `Class.method`.
    static EXPORT_FUNC_MAP: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    // Monotonic counter for synthetic result temporaries of value-producing
    // structured control (block/if with a `(result …)` type).
    static WAST_TEMP_COUNTER: RefCell<usize> = const { RefCell::new(0) };
    // Exception tag name (without `$`) → payload arity, from `(tag $e (param …))`.
    // A `catch $e` needs the arity to bind the right number of payload values.
    static TAG_ARITIES: RefCell<HashMap<String, u8>> = RefCell::new(HashMap::new());
}

/// Payload arity of exception tag `name` (0 if undeclared).
fn tag_arity(name: &str) -> u8 {
    TAG_ARITIES.with(|t| t.borrow().get(name).copied().unwrap_or(0))
}

/// A fresh unique identifier for a structured-control result temporary.
fn fresh_result_temp() -> String {
    WAST_TEMP_COUNTER.with(|c| {
        let mut n = c.borrow_mut();
        let name = format!("__wat_res{}", *n);
        *n += 1;
        name
    })
}

/// Does an unfolded `loop` opener declare a `(param …)` block type? Loop
/// parameters thread stack values across iterations, which the while(true)
/// lowering doesn't model — such loops are emitted once (not looped) so they
/// fail cleanly instead of spinning forever.
fn peek_opener_has_param(pair: &Pair<Rule>) -> bool {
    let inner = if pair.as_rule() == Rule::instr {
        match pair.clone().into_inner().next() {
            Some(p) => p,
            None => return false }
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return false;
    }
    inner.into_inner().any(|c| {
        c.as_rule() == Rule::instr_arg
            && c.clone().into_inner().next().map(|i| i.as_rule()) == Some(Rule::block_type)
            && c.into_inner().next().map(|i| i.as_str().contains("param")) == Some(true)
    })
}

/// How many stack values an unfolded `block`/`loop` opener consumes as block
/// parameters — the total `val_type` count across its `(param …)` block-type
/// immediates. WASM `block (param t*)` pops `t*` off the enclosing stack into
/// the block body; this count lets the fold seed the body with those values
/// instead of discarding them.
fn peek_block_param_count(pair: &Pair<Rule>) -> usize {
    let inner = if pair.as_rule() == Rule::instr {
        match pair.clone().into_inner().next() {
            Some(p) => p,
            None => return 0 }
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return 0;
    }
    let mut count = 0;
    for c in inner.into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        if let Some(bt) = c.into_inner().next() {
            if bt.as_rule() == Rule::block_type && bt.as_str().trim_start().starts_with("(param") {
                count += bt
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::val_type)
                    .count();
            }
        }
    }
    count
}

/// How many result values a `block`/`loop`/`if`/`try` opener yields — the total
/// `any_val_type` count across its `(result …)` block-type immediates, plus any
/// `(type $sig)` immediate's result count. 0 = void, 1 = single-value baseline,
/// N = WASM multi-value.
fn peek_block_result_count(pair: &Pair<Rule>) -> usize {
    let inner = if pair.as_rule() == Rule::instr {
        match pair.clone().into_inner().next() {
            Some(p) => p,
            None => return 0 }
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return 0;
    }
    let mut count = 0;
    for c in inner.into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        let Some(bt) = c.into_inner().next() else {
            continue;
        };
        if bt.as_rule() != Rule::block_type {
            continue;
        }
        let s = bt.as_str().trim_start();
        if s.starts_with("(result") {
            count += bt
                .into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::any_val_type | Rule::val_type))
                .count();
        } else if s.starts_with("(type") {
            // A signature by reference — `(type $sig)` — contributes its
            // declared result count (looked up from the pre-scan).
            if let Some(idx) = bt.into_inner().find(|p| p.as_rule() == Rule::index) {
                let name = idx.as_str().trim_start_matches('$').to_string();
                count += TYPE_FUNC_RESULTS
                    .with(|m| m.borrow().get(&name).copied())
                    .unwrap_or(0);
            }
        }
    }
    count
}

/// Capture a branch/block body's trailing N stack values (the flushed
/// value-statements) into `temps`, `temps[k]` ← the k-th value in stack order
/// (bottom-to-top). The trailing `StmtKind::Expr` run at the end of `body` is
/// exactly the leftover stack the fold flushed; we rewrite each into an
/// assignment.
fn assign_last_n_exprs_to(body: &mut [Statement], temps: &[String]) {
    let n = temps.len();
    if n == 0 {
        return;
    }
    // Indices of the trailing contiguous Expr statements (newest first).
    let mut idxs: Vec<usize> = Vec::with_capacity(n);
    for (i, s) in body.iter().enumerate().rev() {
        if matches!(s.kind, StmtKind::Expr(_)) {
            idxs.push(i);
            if idxs.len() == n {
                break;
            }
        } else {
            break;
        }
    }
    idxs.reverse(); // ascending = stack bottom-to-top
    for (k, &idx) in idxs.iter().enumerate() {
        if let StmtKind::Expr(e) = &body[idx].kind {
            let value = e.clone();
            body[idx].kind = StmtKind::Expr(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident(&temps[k])),
                value: Box::new(value) }));
        }
    }
}

/// Carry the top `temps.len()` stack values into `temps` (temps[0] ← the
/// deepest of the N, temps[last] ← TOS), emitting the assignments into `out`.
/// `consume` pops them (unconditional `br`); otherwise they are peeked (the
/// value passes through a conditional `br_if`). For `temps.len() == 1` this
/// matches the old single-`result_temp` pop/peek behavior.
fn carry_stack_into_temps(
    temps: &[String],
    stack: &mut Vec<Expression>,
    consume: bool,
    out: &mut Vec<Statement>,
) {
    let n = temps.len();
    if n == 0 {
        return;
    }
    let avail = n.min(stack.len());
    let start = stack.len() - avail;
    for (k, temp) in temps.iter().enumerate() {
        if let Some(val) = stack.get(start + k) {
            out.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::ident(temp)),
                    value: Box::new(val.clone()) },
            ))));
        }
    }
    if consume {
        stack.truncate(start);
    }
}

/// Materialize the pending live stack values into temps (in order) and leave
/// those temps on the stack in their place. Used before emitting a block/loop/if
/// STATEMENT: draining live values to bare `Expr` statements would lose their
/// value and stack position (so a later `i32.add` / block-result capture reads
/// the wrong operand); binding them to temps preserves BOTH their side-effect
/// order and their value across the statement boundary.
fn preserve_stack_across_block(stack: &mut Vec<Expression>, statements: &mut Vec<Statement>) {
    let pending: Vec<Expression> = stack.drain(..).collect();
    for e in pending {
        // A bare value expression (const/local.get/ident) has no side effect and
        // needs no temp — keep it deferred as-is. Anything else is bound to a
        // temp so its effect runs here, in order.
        let keep = matches!(e.kind, ExprKind::Lit(_) | ExprKind::Ident(_));
        if keep {
            stack.push(e);
        } else {
            let t = fresh_result_temp();
            statements.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(t.clone()),
                    type_hint: None,
                    init: Some(e),
                    array_bounds: None,
                    with_events: false }],
                kind: VarDeclKind::Let }));
            stack.push(Expression::ident(&t));
        }
    }
}

/// Lower a TOP-LEVEL folded `(block …)`/`(loop …)` as a STATEMENT (mirrors the
/// unfolded block handler), so its body actually runs — `walk_folded_core`'s
/// block path only returned the trailing expression and DISCARDED the body,
/// which silently dropped side effects and any `br`/`br_on_*` inside. Leaves the
/// block's N result values on `stack` for the continuation.
fn emit_folded_block(
    inner: Pair<Rule>,
    is_loop: bool,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let mut label: Option<String> = None;
    let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
    let mut result_count = 0usize;
    let mut param_count = 0usize;
    for child in inner.into_inner() {
        match child.as_rule() {
            Rule::id => label = Some(child.as_str()[1..].to_string()),
            Rule::block_type => {
                let s = child.as_str().trim_start();
                if s.starts_with("(result") {
                    result_count += child
                        .into_inner()
                        .filter(|p| matches!(p.as_rule(), Rule::any_val_type | Rule::val_type))
                        .count();
                } else if s.starts_with("(param") {
                    param_count += child
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::val_type)
                        .count();
                }
            }
            Rule::instr => instr_pairs.push(child),
            _ => {}
        }
    }
    // Seed `(param …)` inputs from the enclosing stack, then flush pending side
    // effects (mirrors the unfolded block handler).
    let seed = if param_count > 0 && stack.len() >= param_count {
        stack.split_off(stack.len() - param_count)
    } else {
        Vec::new()
    };
    preserve_stack_across_block(stack, statements);
    let result_temps: Vec<String> = (0..result_count).map(|_| fresh_result_temp()).collect();
    for tmp in &result_temps {
        statements.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp.clone()),
                type_hint: None,
                init: Some(Expression::null()),
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Let }));
    }
    let kind = if is_loop {
        LabelKind::Loop
    } else {
        LabelKind::Block
    };
    let effective = labels.push(label.clone(), kind, result_temps.clone());
    let mut body = fold_instructions_seeded(instr_pairs, labels, seed)?;
    labels.pop();
    assign_last_n_exprs_to(&mut body, &result_temps);
    let inner_stmt = if !is_loop {
        Statement::with_span(StmtKind::Block(body), span)
    } else {
        body.push(Statement::with_span(
            StmtKind::Break(BreakTarget::Implicit),
            span,
        ));
        Statement::with_span(
            StmtKind::While {
                cond: Expression::bool(true),
                body,
                else_body: None },
            span,
        )
    };
    statements.push(Statement::with_span(
        StmtKind::Labeled {
            label: effective,
            body: Box::new(inner_stmt) },
        span,
    ));
    for tmp in &result_temps {
        stack.push(Expression::ident(tmp));
    }
    Ok(())
}

/// Lower a canonical folded `(try_table (catch $tag $L) (catch_ref $tag $L)
/// (catch_all $L) (catch_all_ref $L) body…)` (WASM 3.0 exception handling).
/// Each clause transfers a matching thrown exception to the enclosing label
/// `$L`, delivering the tag's payload — and, for `_ref` clauses, the caught
/// `exnref` — exactly like a `br $L` carrying those values. Reuses the inline
/// `WasmTryTable` AST: each clause becomes a `WasmCatch` whose handler carries
/// the delivered payload/exnref into `$L`'s branch-carry temps and branches
/// there. The protected body runs normally when nothing is thrown.
fn emit_folded_try_table(
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let mut clauses: Vec<Pair<Rule>> = Vec::new();
    let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
    let mut result_count = 0usize;
    for child in inner.into_inner() {
        match child.as_rule() {
            Rule::try_clause => clauses.push(child),
            Rule::instr => instr_pairs.push(child),
            Rule::block_type => {
                if child.as_str().trim_start().starts_with("(result") {
                    result_count += child
                        .into_inner()
                        .filter(|p| matches!(p.as_rule(), Rule::any_val_type | Rule::val_type))
                        .count();
                }
            }
            _ => {} // id — try_table's own label is unused (clauses target outer blocks)
        }
    }

    // Side effects pending on the stack must run before the protected region.
    preserve_stack_across_block(stack, statements);

    // On NORMAL completion (nothing thrown) the body's trailing values are the
    // try_table's results, captured in temps left on the stack afterwards.
    let result_temps: Vec<String> = (0..result_count).map(|_| fresh_result_temp()).collect();
    for tmp in &result_temps {
        statements.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp.clone()),
                type_hint: None,
                init: Some(Expression::null()),
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Let }));
    }
    let mut body = fold_instructions(instr_pairs, labels)?;
    assign_last_n_exprs_to(&mut body, &result_temps);

    let mut wasm_catches: Vec<WasmCatch> = Vec::new();
    for clause in clauses {
        let kw = clause.as_str();
        let idxs: Vec<String> = clause
            .into_inner()
            .filter(|c| c.as_rule() == Rule::index)
            .map(|c| c.as_str().trim_start_matches('$').to_string())
            .collect();
        // catch/catch_ref: [tag, label]; catch_all/catch_all_ref: [label].
        let (tag, capture_ref, label) = if kw.starts_with("(catch_all_ref") {
            (None, true, idxs.first().cloned().unwrap_or_default())
        } else if kw.starts_with("(catch_all") {
            (None, false, idxs.first().cloned().unwrap_or_default())
        } else if kw.starts_with("(catch_ref") {
            (
                idxs.first().cloned(),
                true,
                idxs.get(1).cloned().unwrap_or_default(),
            )
        } else {
            (
                idxs.first().cloned(),
                false,
                idxs.get(1).cloned().unwrap_or_default(),
            )
        };

        let arity = tag.as_deref().map(tag_arity).unwrap_or(0);
        let payload_binds: Vec<String> = (0..arity).map(|_| fresh_result_temp()).collect();
        let exnref_bind = if capture_ref {
            Some(fresh_result_temp())
        } else {
            None
        };

        // Handler ≡ `br $L` carrying the delivered payload (+exnref): assign them
        // into the target's carry temps, then branch. The compiler binds the
        // caught payload/exnref into these locals before running this body.
        let entry = labels
            .resolve(&BrTarget::Named(label.clone()))
            .ok_or_else(|| format!("try_table clause targets unknown label ${label}"))?;
        let carry = branch_carry_temps(&entry);
        let mut hstack: Vec<Expression> =
            payload_binds.iter().map(|n| Expression::ident(n)).collect();
        if let Some(e) = &exnref_bind {
            hstack.push(Expression::ident(e));
        }
        let mut hbody: Vec<Statement> = Vec::new();
        carry_stack_into_temps(&carry, &mut hstack, true, &mut hbody);
        hbody.push(br_stmt_for(&entry, span));

        wasm_catches.push(WasmCatch {
            tag,
            payload_binds,
            capture_ref,
            exnref_bind,
            body: hbody });
    }

    statements.push(Statement::with_span(
        StmtKind::WasmTryTable {
            body,
            catches: wasm_catches },
        span,
    ));
    // Normal-completion results are now available to the enclosing context.
    for tmp in &result_temps {
        stack.push(Expression::ident(tmp));
    }
    Ok(())
}

/// Lower a folded `(br_on_null $L operand)` / `(br_on_non_null $L operand)` as a
/// structured conditional branch (the VM opcode uses a raw ip-offset that does
/// not fit the walker's Break model). `br_on_null` branches when the ref IS
/// null (carrying the values below it); the non-null ref stays on the stack for
/// fall-through. `br_on_non_null` branches when the ref is NON-null (carrying
/// the ref into the target's result); the null case drops the ref.
fn emit_folded_br_on_null(
    inner: Pair<Rule>,
    is_non_null: bool,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    // Operands and the label both arrive as `instr_arg` (a nested folded
    // operand matches `instr_arg → folded_instr`; the `$L` label is `instr_arg
    // → id`). The label is the id/index arg; everything else is a value operand
    // folded onto the stack (the ref ends up on top).
    let mut label_arg: Option<Expression> = None;
    for child in inner.into_inner() {
        match child.as_rule() {
            Rule::instr_arg => {
                let is_label = matches!(
                    child.clone().into_inner().next().map(|x| x.as_rule()),
                    Some(Rule::id) | Some(Rule::index)
                );
                if is_label && label_arg.is_none() {
                    label_arg = Some(walk_instr_arg_pair(child, labels)?);
                } else {
                    let e = walk_instr_arg_pair(child, labels)?;
                    stack.push(e);
                }
            }
            Rule::instr => {
                let e = walk_instr_as_expr(child, labels)?;
                stack.push(e);
            }
            _ => {}
        }
    }
    let ref_val = stack.pop().unwrap_or_else(Expression::null);
    let tmp = fresh_result_temp();
    statements.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(tmp.clone()),
            type_hint: None,
            init: Some(ref_val),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Let }));
    let is_null = make_call("ref_is_null", vec![Expression::ident(&tmp)], span);
    let target = br_target_of(label_arg.as_ref());
    let mut then_body: Vec<Statement> = Vec::new();
    match labels.resolve(&target) {
        Some(entry) => {
            if is_non_null {
                // Carry the ref into the target's topmost result, then branch.
                if let Some(rt) = entry.result_temps.last() {
                    then_body.push(Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Assign {
                            target: Box::new(Expression::ident(rt)),
                            value: Box::new(Expression::ident(&tmp)) },
                    ))));
                }
            } else {
                // br_on_null carries the values BELOW the ref (peeked).
                carry_stack_into_temps(&entry.result_temps, stack, false, &mut then_body);
            }
            then_body.push(br_stmt_for(&entry, span));
        }
        None => then_body.push(make_br_stmt_opt(None, labels, span)) }
    let cond = if is_non_null {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(is_null),
            right: Box::new(Expression::int(0)) })
    } else {
        is_null
    };
    statements.push(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body: None,
            elifs: Vec::new() },
        span,
    ));
    // Fall-through: br_on_null leaves the (non-null) ref; br_on_non_null drops it.
    if !is_non_null {
        stack.push(Expression::ident(&tmp));
    }
    Ok(())
}

/// Lower a folded `(return operand*)` as a `return` statement (the generic path
/// maps `return` to a null expression, losing both the branch and its value).
/// Multi-value-aware, like the plain `return` handler.
fn emit_folded_return(
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    for child in inner.into_inner() {
        match child.as_rule() {
            // A folded return value nests as `instr_arg → folded_instr`.
            Rule::instr_arg => stack.push(walk_instr_arg_pair(child, labels)?),
            Rule::instr => stack.push(walk_instr_as_expr(child, labels)?),
            _ => {}
        }
    }
    let n = CURRENT_FN_RESULTS.with(|c| *c.borrow());
    if n >= 2 {
        statements.push(multi_value_return_stmt(stack, n, span));
    } else {
        let val = stack.pop();
        statements.push(Statement::with_span(StmtKind::Return(val), span));
    }
    Ok(())
}

// ── Label context ─────────────────────────────────────────────────────────────
// `br $label` targets a block (Break) or a loop (Continue).  We track which
// as we walk block/loop constructs.

#[derive(Clone, PartialEq)]
enum LabelKind {
    Block,
    Loop }

#[derive(Clone)]
struct LabelEntry {
    /// Always present — a synthetic name is minted when the source omits one, so
    /// every block/loop is addressable (numeric `br N` needs no source label).
    name: String,
    kind: LabelKind,
    /// The result temporaries for a value-producing block/loop — one per result
    /// value (empty = void, len 1 = the single-value baseline, len N = WASM
    /// multi-value). `br` to this frame carries the top N stack values into
    /// them (temps[0] ← deepest of the N, matching stack order).
    result_temps: Vec<String>,
    /// The parameter temporaries for a `loop (param …)` — the synthetic locals
    /// that thread the loop's operand-stack params across iterations. A `br` to
    /// a loop (a `continue`) carries the top N stack values into these before
    /// looping, and the loop body reads them as its seed. Empty for blocks and
    /// param-less loops.
    param_temps: Vec<String> }

/// A fresh synthetic block/loop label.
fn fresh_block_label() -> String {
    WAST_TEMP_COUNTER.with(|c| {
        let mut n = c.borrow_mut();
        let name = format!("__wat_lbl{}", *n);
        *n += 1;
        name
    })
}

struct LabelStack(Vec<LabelEntry>);

impl LabelStack {
    fn new() -> Self {
        LabelStack(Vec::new())
    }
    /// Push a frame, minting a synthetic name if the source has none. Returns the
    /// effective label so the caller can build the matching `Labeled` statement.
    fn push(&mut self, name: Option<String>, kind: LabelKind, result_temps: Vec<String>) -> String {
        let effective = name.unwrap_or_else(fresh_block_label);
        self.0.push(LabelEntry {
            name: effective.clone(),
            kind,
            result_temps,
            param_temps: Vec::new() });
        effective
    }
    fn pop(&mut self) {
        self.0.pop();
    }

    /// Attach loop-parameter temporaries to the just-pushed frame (a
    /// `loop (param …)`), so a `br` back to it threads the next iteration's
    /// param values through them.
    fn set_last_param_temps(&mut self, param_temps: Vec<String>) {
        if let Some(last) = self.0.last_mut() {
            last.param_temps = param_temps;
        }
    }

    fn kind_of(&self, label: &str) -> Option<LabelKind> {
        self.0
            .iter()
            .rev()
            .find(|e| e.name == label)
            .map(|e| e.kind.clone())
    }

    /// Resolve a `br` target: symbolic `$name`, numeric index (0 = innermost), or
    /// None (defaults to innermost).
    fn resolve(&self, target: &BrTarget) -> Option<LabelEntry> {
        match target {
            BrTarget::Named(n) => self.0.iter().rev().find(|e| &e.name == n).cloned(),
            BrTarget::Index(i) => {
                let len = self.0.len();
                (*i < len).then(|| self.0[len - 1 - i].clone())
            }
            BrTarget::Innermost => self.0.last().cloned() }
    }
}

/// How a `br`/`br_if` names its destination frame.
enum BrTarget {
    Named(String),
    Index(usize),
    Innermost }

/// Derive a `br` target from its first argument (label id or numeric index).
fn br_target_of(arg: Option<&Expression>) -> BrTarget {
    match arg.map(|a| &a.kind) {
        Some(ExprKind::Ident(n)) => BrTarget::Named(n.clone()),
        Some(ExprKind::Lit(Literal::Int(i))) => BrTarget::Index(*i as usize),
        _ => BrTarget::Innermost }
}

/// The temporaries a `br`/`br_if` to `entry` carries the top-of-stack values
/// into: a `loop (param …)` continue threads the next iteration's params, so it
/// carries the loop's param temps; every other branch carries the target's
/// result temps.
fn branch_carry_temps(entry: &LabelEntry) -> Vec<String> {
    if entry.kind == LabelKind::Loop && !entry.param_temps.is_empty() {
        entry.param_temps.clone()
    } else {
        entry.result_temps.clone()
    }
}

/// The break/continue statement for a resolved `br` target frame.
fn br_stmt_for(entry: &LabelEntry, span: Span) -> Statement {
    match entry.kind {
        LabelKind::Loop => Statement::with_span(
            StmtKind::Continue(ContinueTarget::Label(entry.name.clone())),
            span,
        ),
        LabelKind::Block => Statement::with_span(
            StmtKind::Break(BreakTarget::Label(entry.name.clone())),
            span,
        ) }
}

// ── Entry point ──────────────────────────────────────────────────────────────

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs =
        WastParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for cmd in top.into_inner() {
                    if cmd.as_rule() != Rule::EOI {
                        walk_script_cmd(cmd, &mut body)?;
                    }
                }
            }
            Rule::EOI => {}
            _ => walk_script_cmd(top, &mut body)? }
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
        directives: Default::default() })
}

// ── Script commands ───────────────────────────────────────────────────────────

fn walk_script_cmd(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::script_cmd => {
            let inner = pair.into_inner().next().ok_or("Empty script_cmd")?;
            walk_script_cmd(inner, body)
        }
        Rule::module => {
            body.extend(walk_module(pair)?);
            Ok(())
        }
        // `(module quote "…")` defers a module given as WAT *text*: unquote the
        // string pieces, concatenate, and parse them as a real module.
        Rule::module_quote_cmd => {
            let text: String = pair
                .into_inner()
                .filter(|c| c.as_rule() == Rule::string)
                .map(|s| unquote(s.as_str()))
                .collect();
            body.extend(parse(&text)?.body);
            Ok(())
        }
        // `(module binary "…")` embeds a module as raw bytes. Decoding the binary
        // section stream is a separate facility; accept and skip it here so the
        // surrounding script still parses (it carries no source-level module).
        Rule::module_binary_cmd => Ok(()),
        Rule::assert_return => {
            body.push(walk_assert_return(pair)?);
            Ok(())
        }
        Rule::assert_trap | Rule::assert_instantiation_trap => {
            body.push(walk_assert_trap(pair)?);
            Ok(())
        }
        Rule::assert_invalid => {
            body.push(walk_assert_generic(pair, "__wast_assert_invalid")?);
            Ok(())
        }
        Rule::assert_malformed => {
            body.push(walk_assert_generic(pair, "__wast_assert_malformed")?);
            Ok(())
        }
        Rule::assert_unlinkable => {
            body.push(walk_assert_generic(pair, "__wast_assert_unlinkable")?);
            Ok(())
        }
        Rule::assert_exhaustion => {
            body.push(walk_assert_generic(pair, "__wast_assert_exhaustion")?);
            Ok(())
        }
        Rule::assert_suspension => {
            body.push(walk_assert_generic(pair, "__wast_assert_suspension")?);
            Ok(())
        }
        Rule::invoke_cmd => {
            body.push(walk_invoke_cmd(pair)?);
            Ok(())
        }
        Rule::register_cmd => {
            body.push(walk_register_cmd(pair)?);
            Ok(())
        }
        Rule::get_cmd => {
            body.push(walk_get_cmd(pair)?);
            Ok(())
        }
        _ => Ok(()) }
}

// ── Module ────────────────────────────────────────────────────────────────────

/// Recursively collect the `$id` targets of every `global.set` instruction in
/// a subtree (used to catch writes to immutable globals during validation).
fn collect_global_set_targets(pair: Pair<Rule>, out: &mut Vec<String>) {
    let is_set = matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr)
        && pair
            .clone()
            .into_inner()
            .any(|c| c.as_rule() == Rule::instr_name && c.as_str() == "global.set");
    if is_set {
        // The written global is the first `id`/`index` immediate.
        for arg in pair.clone().into_inner() {
            if arg.as_rule() == Rule::instr_arg {
                if let Some(id) = arg.into_inner().find(|c| c.as_rule() == Rule::id) {
                    out.push(id.as_str()[1..].to_string());
                    break;
                }
            }
        }
    }
    for child in pair.into_inner() {
        collect_global_set_targets(child, out);
    }
}

/// WASM validation checks that must reject a module at parse time (the
/// `parse_err` spec tests): duplicate export names, a `start` referencing an
/// undefined function, and a `global.set` on an immutable global. Returns the
/// first violation found.
fn validate_module(pair: &Pair<Rule>) -> Result<(), String> {
    use std::collections::HashSet;
    let mut export_names: HashSet<String> = HashSet::new();
    let mut func_names: HashSet<String> = HashSet::new();
    let mut func_count: usize = 0;
    let mut immut_globals: HashSet<String> = HashSet::new();
    let mut start_target: Option<String> = None;
    // WASM 3.0 §6.4: imports occupy the low end of each index space, so the text
    // format requires every import to precede all non-import definitions. An
    // `(import …)` (or an inline `(func (import …))` etc.) after a real func/
    // table/memory/global/tag definition is a well-formedness error.
    let mut def_seen = false;

    for field in pair.clone().into_inner() {
        if field.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = field.into_inner().next() else {
            continue;
        };
        let is_def_kind = matches!(
            inner.as_rule(),
            Rule::func_field
                | Rule::table_field
                | Rule::memory_field
                | Rule::global_field
                | Rule::tag_field
        );
        let is_inline_import = is_def_kind && inner.as_str().contains("(import");
        let is_import = inner.as_rule() == Rule::import_field || is_inline_import;
        if is_import && def_seen {
            return Err("imports must occur before all non-import definitions".to_string());
        }
        if is_def_kind && !is_inline_import {
            def_seen = true;
        }
        match inner.as_rule() {
            Rule::export_field => {
                if let Some(name) = inner
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::string)
                    .map(|s| unquote(s.as_str()))
                {
                    if !export_names.insert(name.clone()) {
                        return Err(format!("duplicate export name: \"{}\"", name));
                    }
                }
            }
            Rule::func_field => {
                func_count += 1;
                if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                    func_names.insert(id.as_str()[1..].to_string());
                }
            }
            Rule::import_field => {
                // An imported func also participates in `start` resolution.
                let children: Vec<_> = inner.into_inner().collect();
                if let Some(desc) = children.iter().find(|c| c.as_rule() == Rule::import_desc) {
                    let dtext = desc.as_str();
                    if dtext.trim_start().starts_with("(func") || dtext.contains("(func") {
                        func_count += 1;
                        if let Some(id) =
                            desc.clone().into_inner().find(|c| c.as_rule() == Rule::id)
                        {
                            func_names.insert(id.as_str()[1..].to_string());
                        }
                    }
                }
            }
            Rule::global_field => {
                let children: Vec<_> = inner.into_inner().collect();
                let id = children
                    .iter()
                    .find(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
                let is_mut = children
                    .iter()
                    .any(|c| c.as_rule() == Rule::global_type && c.as_str().contains("mut"));
                if let Some(id) = id {
                    if !is_mut {
                        immut_globals.insert(id);
                    }
                }
            }
            Rule::start_field => {
                if let Some(idx) = inner.into_inner().find(|c| c.as_rule() == Rule::index) {
                    start_target = Some(idx.as_str().to_string());
                }
            }
            _ => {}
        }
    }

    if let Some(t) = start_target {
        if let Some(name) = t.strip_prefix('$') {
            if !func_names.contains(name) {
                return Err(format!("unknown start function: {}", t));
            }
        } else if let Ok(n) = t.parse::<usize>() {
            if n >= func_count {
                return Err(format!("unknown start function index: {}", n));
            }
        }
    }

    if !immut_globals.is_empty() {
        let mut targets = Vec::new();
        collect_global_set_targets(pair.clone(), &mut targets);
        for t in targets {
            if immut_globals.contains(&t) {
                return Err(format!("global.set on immutable global: ${}", t));
            }
        }
    }

    Ok(())
}

fn walk_module(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    validate_module(&pair)?;
    let span = to_span(&pair);
    let mut module_name: Option<String> = None;
    let mut members: Vec<ClassMember> = Vec::new();
    let mut pre_stmts: Vec<Statement> = Vec::new(); // before class (globals)
    let mut post_stmts: Vec<Statement> = Vec::new(); // after class (start, exports, imports)

    let mut index_arities = Vec::new();
    let mut name_arities = HashMap::new();
    let mut index_results = Vec::new();
    let mut name_results = HashMap::new();

    // 1. Pre-scan imports. Params live inside `typeuse` (and, for imports, inside
    //    `import_desc`), so the signature scan must descend, not read direct children.
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::import_field {
                    let (name, params_count, results_count) = scan_func_signature(inner);
                    index_arities.push(params_count);
                    index_results.push(results_count);
                    if let Some(n) = name {
                        name_arities.insert(n.clone(), params_count);
                        name_results.insert(n, results_count);
                    }
                }
            }
        }
    }

    // 2. Pre-scan defined functions
    let mut defined_names: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut export_map: HashMap<String, String> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::func_field {
                    let (name, params_count, results_count) = scan_func_signature(inner.clone());
                    index_arities.push(params_count);
                    index_results.push(results_count);
                    if let Some(n) = &name {
                        defined_names.insert(n.clone());
                        name_arities.insert(n.clone(), params_count);
                        name_results.insert(n.clone(), results_count);
                    }
                    // Inline exports: `(func $id (export "e") …)`. The method name
                    // is the id, or (for an unnamed func) its first export name.
                    let exports: Vec<String> = inner
                        .into_inner()
                        .filter(|c| c.as_rule() == Rule::export_inline)
                        .filter_map(|c| c.into_inner().find(|p| p.as_rule() == Rule::string))
                        .map(|s| unquote(s.as_str()))
                        .collect();
                    let method = name.clone().or_else(|| exports.first().cloned());
                    if let Some(m) = method {
                        // An unnamed exported func is reached by its export name
                        // (e.g. `_start`); key its signature there too so the
                        // entry auto-invoke can tell whether it yields a value.
                        name_results.entry(m.clone()).or_insert(results_count);
                        name_arities.entry(m.clone()).or_insert(params_count);
                        for e in exports {
                            export_map.insert(e, m.clone());
                        }
                    }
                } else if inner.as_rule() == Rule::export_field {
                    // `(export "e" (func $g))`: map the export name to the func id.
                    let mut ename: Option<String> = None;
                    let mut target: Option<String> = None;
                    for c in inner.into_inner() {
                        match c.as_rule() {
                            Rule::string => ename = Some(unquote(c.as_str())),
                            Rule::id => target = Some(c.as_str()[1..].to_string()),
                            _ => {}
                        }
                    }
                    if let (Some(e), Some(t)) = (ename, target) {
                        export_map.insert(e, t);
                    }
                }
            }
        }
    }

    FUNC_INDEX_ARITIES.with(|f| *f.borrow_mut() = index_arities);
    FUNC_NAME_ARITIES.with(|f| *f.borrow_mut() = name_arities);
    FUNC_INDEX_RESULTS.with(|f| *f.borrow_mut() = index_results);
    FUNC_NAME_RESULTS.with(|f| *f.borrow_mut() = name_results);
    DEFINED_FUNC_NAMES.with(|f| *f.borrow_mut() = defined_names);
    EXPORT_FUNC_MAP.with(|f| *f.borrow_mut() = export_map);

    // 3. Pre-scan struct type definitions to know field counts for struct.new arity
    let mut struct_counts: HashMap<String, usize> = HashMap::new();
    let mut func_param_counts: HashMap<String, usize> = HashMap::new();
    let mut func_result_counts: HashMap<String, usize> = HashMap::new();
    let mut array_elem_types: HashMap<String, String> = HashMap::new();
    // Every declared type's name, in order, so a numeric parent index resolves
    // to a name. GC structs collected here (name, raw parent ref, field count).
    let mut type_order: Vec<String> = Vec::new();
    let mut struct_types_raw: Vec<(String, Option<String>, usize)> = Vec::new();
    let mut struct_field_types_map: HashMap<String, Vec<String>> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::type_field {
                    let mut type_name: Option<String> = None;
                    let mut field_count = 0usize;
                    let mut field_types: Vec<String> = Vec::new();
                    let mut is_struct = false;
                    let mut func_params: Option<usize> = None;
                    let mut func_results: Option<usize> = None;
                    let mut array_elem: Option<String> = None;
                    // Parent type reference (`$Base` or a numeric index) from a
                    // `struct_subtype`/`array_subtype` trailing index or a
                    // `(sub $Base …)` leading index. Resolved to a name below.
                    let mut parent_ref: Option<String> = None;
                    for sub in inner.into_inner() {
                        match sub.as_rule() {
                            Rule::id => type_name = Some(sub.as_str()[1..].to_string()),
                            // `(sub final? $super* composite)` — the standard GC
                            // subtype wrapper: capture the first supertype, then
                            // fall through to its inner composite type.
                            Rule::sub_type => {
                                let mut composite = None;
                                for c in sub.into_inner() {
                                    match c.as_rule() {
                                        Rule::index => {
                                            if parent_ref.is_none() {
                                                parent_ref = Some(
                                                    c.as_str().trim_start_matches('$').to_string(),
                                                );
                                            }
                                        }
                                        Rule::composite_type => composite = Some(c),
                                        _ => {}
                                    }
                                }
                                if let Some(inner2) = composite.and_then(|c| c.into_inner().next())
                                {
                                    if inner2.as_rule() == Rule::array_type {
                                        array_elem = array_elem_type(&inner2);
                                    }
                                    if matches!(
                                        inner2.as_rule(),
                                        Rule::struct_type | Rule::struct_subtype
                                    ) {
                                        is_struct = true;
                                        field_types = struct_field_types(&inner2);
                                        field_count = field_types.len();
                                    }
                                }
                            }
                            Rule::composite_type => {
                                if let Some(inner2) = sub.into_inner().next() {
                                    if inner2.as_rule() == Rule::array_type {
                                        array_elem = array_elem_type(&inner2);
                                    }
                                    match inner2.as_rule() {
                                        Rule::struct_type => {
                                            is_struct = true;
                                            field_types = struct_field_types(&inner2);
                                            field_count = field_types.len();
                                        }
                                        // `(struct_subtype field* $Base)` — legacy
                                        // GC-MVP form: fields then the supertype.
                                        Rule::struct_subtype => {
                                            is_struct = true;
                                            field_types = struct_field_types(&inner2);
                                            field_count = field_types.len();
                                            if let Some(idx) = inner2
                                                .into_inner()
                                                .filter(|p| p.as_rule() == Rule::index)
                                                .next_back()
                                            {
                                                parent_ref = Some(
                                                    idx.as_str()
                                                        .trim_start_matches('$')
                                                        .to_string(),
                                                );
                                            }
                                        }
                                        Rule::func_type => {
                                            // param / result count = total val types
                                            // across all `(param …)` / `(result …)`.
                                            let mut ps = 0usize;
                                            let mut rs = 0usize;
                                            for p in inner2.into_inner() {
                                                let n = p
                                                    .clone()
                                                    .into_inner()
                                                    .filter(|v| v.as_rule() == Rule::any_val_type)
                                                    .count();
                                                match p.as_rule() {
                                                    Rule::param => ps += n,
                                                    Rule::result => rs += n,
                                                    _ => {}
                                                }
                                            }
                                            func_params = Some(ps);
                                            func_results = Some(rs);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    // A type declared without an id is referenced only by index;
                    // give it a stable synthetic name so index→name still works.
                    let name = type_name
                        .clone()
                        .unwrap_or_else(|| format!("__wast_type_{}", type_order.len()));
                    type_order.push(name.clone());
                    if is_struct {
                        struct_counts.insert(name.clone(), field_count);
                        struct_types_raw.push((name.clone(), parent_ref.clone(), field_count));
                        struct_field_types_map.insert(name.clone(), field_types.clone());
                    }
                    if let Some(r) = func_results {
                        func_result_counts.insert(name.clone(), r);
                    }
                    if let Some(n) = func_params {
                        func_param_counts.insert(name.clone(), n);
                    }
                    if let Some(e) = array_elem {
                        array_elem_types.insert(name.clone(), e);
                    }
                }
            }
        }
    }
    // Resolve each struct's parent reference (a `$name` kept verbatim, or a
    // numeric index mapped through declaration order) to a concrete type name.
    let struct_types: Vec<(String, Option<String>, usize)> = struct_types_raw
        .into_iter()
        .map(|(name, parent_ref, fields)| {
            let parent = parent_ref.and_then(|p| {
                if let Ok(i) = p.parse::<usize>() {
                    type_order.get(i).cloned()
                } else {
                    Some(p)
                }
            });
            (name, parent, fields)
        })
        .collect();
    STRUCT_TYPES.with(|f| *f.borrow_mut() = struct_types);
    STRUCT_FIELD_TYPES.with(|f| *f.borrow_mut() = struct_field_types_map);
    STRUCT_FIELD_COUNTS.with(|f| *f.borrow_mut() = struct_counts);
    TYPE_FUNC_PARAMS.with(|f| *f.borrow_mut() = func_param_counts);
    TYPE_FUNC_RESULTS.with(|f| *f.borrow_mut() = func_result_counts);
    ARRAY_ELEM_TYPE.with(|f| *f.borrow_mut() = array_elem_types);
    ELEM_SEG_COUNTER.with(|c| *c.borrow_mut() = 0);

    // 3a. Pre-scan tables so named tables (`$t1`) resolve to their declaration
    //     index for `elem` population and `call_indirect $t` dispatch.
    let mut table_names: HashMap<String, usize> = HashMap::new();
    let mut table_idx = 0usize;
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::table_field {
                    if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                        table_names.insert(id.as_str()[1..].to_string(), table_idx);
                    }
                    table_idx += 1;
                }
            }
        }
    }
    TABLE_NAME_INDEX.with(|f| *f.borrow_mut() = table_names);

    // 3a'. Pre-scan memories so named memories (`$m2`) resolve to their
    //      declaration index for multi-memory load/store/size and data segments.
    let mut memory_names: HashMap<String, usize> = HashMap::new();
    let mut memory_idx = 0usize;
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::memory_field {
                    if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                        memory_names.insert(id.as_str()[1..].to_string(), memory_idx);
                    }
                    memory_idx += 1;
                }
            }
        }
    }
    MEMORY_NAME_INDEX.with(|f| *f.borrow_mut() = memory_names);

    // 3a''. Pre-scan globals so a `global.get N` / `global.set N` by numeric
    //       index resolves to the right binding (each global's `$id`, or a
    //       synthetic `__wasm_global_<i>` when unnamed).
    let mut global_names: Vec<String> = Vec::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::global_field {
                    let idx = global_names.len();
                    global_names.push(global_binding_name(&inner, idx));
                }
            }
        }
    }
    GLOBAL_INDEX_NAME.with(|f| *f.borrow_mut() = global_names);

    // 3b. Pre-scan exception tags so a `catch $e` in any function body knows
    //     the tag's payload arity regardless of source order. Reset first —
    //     the thread-local persists across modules compiled on this thread.
    let mut tag_arities: HashMap<String, u8> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::tag_field {
                    let (name, arity) = scan_tag_signature(inner);
                    if let Some(name) = name {
                        tag_arities.insert(name, arity);
                    }
                }
            }
        }
    }
    TAG_ARITIES.with(|t| *t.borrow_mut() = tag_arities);

    // 4. Detect the WASI command entry. A module that exports a function as
    //    "_start" is a command module — instantiation runs `_start` with no
    //    driver. Explicit `(start $f)` fields are handled separately below; if
    //    one is present we don't also auto-run `_start`.
    let mut start_export_name: Option<String> = None;
    let mut start_fn_name: Option<String> = None;
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        match inner.as_rule() {
            Rule::start_field => {
                // Capture the start function's `$id` so it can be invoked as a
                // static method of the module class at instantiation.
                start_fn_name = inner
                    .into_inner()
                    .next()
                    .and_then(|idx| idx.into_inner().next())
                    .filter(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
            }
            Rule::func_field => {
                let mut id: Option<String> = None;
                let mut exports_start = false;
                for sub in inner.into_inner() {
                    match sub.as_rule() {
                        Rule::id => id = Some(sub.as_str()[1..].to_string()),
                        Rule::export_inline => {
                            if let Some(s) = sub.into_inner().find(|p| p.as_rule() == Rule::string)
                            {
                                if unquote(s.as_str()) == "_start" {
                                    exports_start = true;
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if exports_start {
                    // walk_func_field names an unnamed exported func after its
                    // export, so the callable name is the id if present else "_start".
                    start_export_name = Some(id.unwrap_or_else(|| "_start".to_string()));
                }
            }
            _ => {}
        }
    }

    // Record the module class name before walking bodies so `call $f` to a
    // defined function can be qualified as `ClassName.f(...)`.
    let prescan_class_name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
        .unwrap_or_else(|| "__wasm_module".to_string());
    MODULE_CLASS_NAME.with(|c| *c.borrow_mut() = prescan_class_name);

    let mut global_decl_idx = 0usize;
    let mut table_decl_idx = 0usize;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {
                module_name = Some(child.as_str()[1..].to_string());
            }
            Rule::module_field => {
                let inner = child.into_inner().next().ok_or("Empty module_field")?;
                match inner.as_rule() {
                    Rule::func_field => {
                        members.push(ClassMember::Method(Box::new(walk_func_field(inner)?)))
                    }
                    Rule::import_field => post_stmts.push(walk_import_field(inner)?),
                    Rule::export_field => {
                        post_stmts.push(Statement::new(StmtKind::Expr(walk_export_field(inner)?)));
                    }
                    Rule::global_field => {
                        // Globals become top-level let bindings BEFORE the class so that
                        // global.get $name → Ident("name") resolves correctly from methods.
                        // The declaration index gives unnamed globals a stable name that
                        // `global.get <idx>` resolves to (see GLOBAL_INDEX_NAME).
                        let (name, init) = walk_global_field(inner, global_decl_idx)?;
                        global_decl_idx += 1;
                        pre_stmts.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(name),
                                type_hint: None,
                                init: Some(init),
                                array_bounds: None,
                                with_events: false }],
                            kind: VarDeclKind::Let }));
                    }
                    // `(start $f)` is invoked as a static method at instantiation
                    // in the module assembly below (its name was captured in the
                    // pre-scan); nothing to emit per-field here.
                    Rule::start_field => {}
                    // Linear memory + data segments: emitted before the class so
                    // the compiler lowers them into the script chunk's memory /
                    // data tables (the VM allocates pages and writes active data
                    // at instantiation, before `_start`).
                    Rule::memory_field => pre_stmts.push(walk_memory_field(inner)?),
                    Rule::data_field => pre_stmts.push(walk_data_field(inner)?),
                    Rule::table_field => {
                        let (decl, population) = walk_table_field(inner, table_decl_idx)?;
                        pre_stmts.push(decl);
                        post_stmts.extend(population);
                        table_decl_idx += 1;
                    }
                    // Exception tags: declared before the class so the tag
                    // entity exists in the script chunk; `throw`/`catch` in the
                    // function chunks re-import by name and coalesce to it.
                    Rule::tag_field => pre_stmts.push(walk_tag_field(inner)?),
                    // Active element segment: populate the funcref table so
                    // call_indirect can dispatch through it. Emitted AFTER the
                    // class (a post-stmt) so the `ref.func` tear-off can resolve
                    // each function's chunk — but still before `_start` runs.
                    Rule::elem_field => post_stmts.push(walk_elem_field(inner)?),
                    _ => {} // type — structural metadata
                }
            }
            _ => {}
        }
    }

    let name = module_name.unwrap_or_else(|| "__wasm_module".to_string());
    let class_name = name.clone();
    let class = Statement::with_span(
        StmtKind::ClassDecl {
            name,
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new() },
        span,
    );

    // Register GC struct types (name, parent, field count) so the compiler
    // installs them in the type table with subtype edges — a compile-time
    // directive (`__wast_register_struct_type`) that emits no runtime code.
    // Emitted first so the identity is known before any `struct.new`/`ref.*`.
    let mut result: Vec<Statement> = STRUCT_TYPES.with(|f| {
        f.borrow()
            .iter()
            .map(|(name, parent, fields)| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_struct_type")),
                    args: vec![
                        Argument::positional(Expression::string(name)),
                        Argument::positional(Expression::string(parent.as_deref().unwrap_or(""))),
                        Argument::positional(Expression::int(*fields as i64)),
                    ],
                    optional: false })))
            })
            .collect()
    });
    // Register GC array types with their element storage type (compile-time
    // directive `__wast_register_array_type(name, elem)`) so the VM can recover
    // the element byte width for `array.init_data`/packed reads. Emitted before
    // the class so the type is element-typed ahead of any `array.*`.
    result.extend(ARRAY_ELEM_TYPE.with(|f| {
        f.borrow()
            .iter()
            .map(|(name, elem)| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_array_type")),
                    args: vec![
                        Argument::positional(Expression::string(name)),
                        Argument::positional(Expression::string(elem)),
                    ],
                    optional: false })))
            })
            .collect::<Vec<_>>()
    }));
    result.extend(pre_stmts);
    result.push(class);
    result.extend(post_stmts);

    // `(start $f)` runs at instantiation: invoke it as a static method of the
    // module class (functions are static methods). This is INDEPENDENT of the
    // `_start` command entry — both run (start first).
    if let Some(sf) = &start_fn_name {
        if start_export_name.as_deref() != Some(sf.as_str()) {
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class_name)),
                    field: sf.clone(),
                    null_safe: false })),
                args: Vec::new(),
                optional: false });
            result.push(Statement::new(StmtKind::Expr(call)));
        }
    }

    // Auto-run the command entry `_start` at instantiation.
    {
        if let Some(entry) = start_export_name {
            // Functions are static methods of the module class, so the entry is
            // reached as `ModuleClass._start()`.
            // Does the entry yield a value? If so, surface it as output the way
            // `wasmtime --invoke` prints an exported function's result.
            let entry_yields = FUNC_NAME_RESULTS
                .with(|f| f.borrow().get(&entry).copied())
                .unwrap_or(0)
                > 0;
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class_name)),
                field: entry,
                null_safe: false });
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args: Vec::new(),
                optional: false });
            let stmt = if entry_yields {
                // `log(entry())` — the entry's declared result is its output.
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("log")),
                    args: vec![Argument::positional(call)],
                    optional: false })
            } else {
                call
            };
            result.push(Statement::new(StmtKind::Expr(stmt)));
        }
    }
    Ok(result)
}

// ── Function field ────────────────────────────────────────────────────────────

fn walk_func_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut result_count: usize = 0;
    let mut body: Vec<Statement> = Vec::new();
    let mut export_names: Vec<String> = Vec::new();
    let mut labels = LabelStack::new();

    let mut instr_pairs = Vec::new();

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
            Rule::import_inline => {}
            Rule::typeuse => {
                params = walk_typeuse_params(child.clone())?;
                // Inline `(result …)` count.
                result_count = child
                    .clone()
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::result)
                    .map(|r| {
                        r.into_inner()
                            .filter(|v| matches!(v.as_rule(), Rule::any_val_type | Rule::val_type))
                            .count()
                    })
                    .sum();
                // A signature given by reference — `(func $f (type $sig) …)` —
                // has no inline params/results; expand the referenced type's
                // shape so `param_count`/`result_arity` (the call_indirect type
                // check) are correct. Placeholder param types suffice: the VM is
                // untyped and the check is over the param/result COUNTS.
                if params.is_empty() && result_count == 0 {
                    if let Some(sig) = child
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::index)
                        .map(|i| i.as_str().trim_start_matches('$').to_string())
                    {
                        let pc = TYPE_FUNC_PARAMS
                            .with(|m| m.borrow().get(&sig).copied())
                            .unwrap_or(0);
                        result_count = TYPE_FUNC_RESULTS
                            .with(|m| m.borrow().get(&sig).copied())
                            .unwrap_or(0);
                        params = (0..pc)
                            .map(|i| Param {
                                name: format!("p{}", i),
                                type_hint: Some("i32".into()),
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false })
                            .collect();
                    }
                }
            }
            Rule::local => {
                body.extend(walk_local(child)?);
            }
            Rule::instr => {
                instr_pairs.push(child);
            }
            _ => {}
        }
    }

    // Expose the function's result count so a `return` inside a multi-value
    // function reraises the top N values as a uniform tuple (multi-value ABI).
    CURRENT_FN_RESULTS.with(|c| *c.borrow_mut() = result_count);
    body.extend(fold_instructions(instr_pairs, &mut labels)?);
    CURRENT_FN_RESULTS.with(|c| *c.borrow_mut() = 0);

    if func_name.is_empty() {
        func_name = export_names
            .first()
            .cloned()
            .unwrap_or_else(|| "__wasm_func".to_string());
    }

    if result_count >= 2 {
        apply_multi_value_return(&mut body, result_count);
    } else {
        apply_implicit_return(&mut body);
    }

    let mut modifiers = Modifiers::default();
    modifiers.is_static = true;

    // Encode the result count in `return_type` (one placeholder type per
    // result) so the compiler can set `chunk.result_arity` — half of the
    // function's type shape for the `call_indirect` runtime check. `None` = a
    // no-result (void) function, distinct from the default 1-value ABI.
    let return_type = if result_count == 0 {
        None
    } else {
        Some(vec!["i32"; result_count].join(","))
    };
    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name: func_name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false },
        span,
    ))
}

/// Recursively read a func/import field's signature: its (first) id and its
/// parameter count. Parameters are wrapped in `typeuse`, and imported funcs are
/// further wrapped in `import_desc`, so a flat scan of direct children misses
/// them — the call-site arity would then be 0 and stack operands never consumed.
fn scan_func_signature(pair: Pair<Rule>) -> (Option<String>, usize, usize) {
    let mut name: Option<String> = None;
    let mut count = 0usize;
    let mut results = 0usize;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {
                if name.is_none() {
                    name = Some(child.as_str()[1..].to_string());
                }
            }
            Rule::param => {
                // `(param $id t)` is one slot; `(param t1 t2 …)` is one per type.
                let mut has_id = false;
                let mut types = 0usize;
                for p in child.into_inner() {
                    match p.as_rule() {
                        Rule::id => has_id = true,
                        // Types are wrapped in `any_val_type` (which may hold a
                        // plain `val_type` or a `(ref …)` form).
                        Rule::any_val_type | Rule::val_type => types += 1,
                        _ => {}
                    }
                }
                count += if has_id { 1 } else { types };
            }
            Rule::result => {
                // `(result t1 t2 …)` yields one value per type.
                results += child
                    .into_inner()
                    .filter(|v| matches!(v.as_rule(), Rule::any_val_type | Rule::val_type))
                    .count();
            }
            Rule::typeuse | Rule::import_desc => {
                let (n, c, r) = scan_func_signature(child);
                if name.is_none() {
                    name = n;
                }
                count += c;
                results += r;
            }
            _ => {}
        }
    }
    (name, count, results)
}

fn walk_typeuse_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::param {
            // `local.get N` indexes params by their ABSOLUTE position across all
            // `(param …)` groups, so auto-name unnamed params `p{running_index}`
            // — not per-group (which made a second `(param i32)` collide on p0).
            let base = params.len();
            params.extend(walk_param(child, base)?);
        }
    }
    Ok(params)
}

fn walk_param(pair: Pair<Rule>, base: usize) -> Result<Vec<Param>, String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::any_val_type | Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }
    if types.is_empty() {
        return Ok(Vec::new());
    }
    if let Some(n) = name {
        return Ok(vec![Param {
            name: n,
            type_hint: types.into_iter().next().map(Into::into),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false }]);
    }
    Ok(types
        .into_iter()
        .enumerate()
        .map(|(i, t)| Param {
            name: format!("p{}", base + i),
            type_hint: Some(t.into()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false })
        .collect())
}

fn walk_local(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::any_val_type | Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }
    Ok(types
        .iter()
        .enumerate()
        .map(|(i, t)| {
            let var_name = name.clone().unwrap_or_else(|| format!("local{}", i));
            let init = match t.as_str() {
                "i32" | "i64" => Expression::int(0),
                // f32 is WASM-exclusive: demote 0.0 to single precision so the
                // default lands as `Value::F32(0.0)` (Displays "0.0"), not a
                // generic float folding to `F64` (Displays "0"). f64 keeps the
                // JS-number 0.0 (Displays "0", matching its shared semantics).
                "f32" => make_call(
                    "f32_demote_f64",
                    vec![Expression::float(0.0)],
                    Span::default(),
                ),
                "f64" => Expression::float(0.0),
                // A concrete `(ref null $t)` local defaults to a WASM GC typed
                // null so `struct.get`/`array.get` on a never-assigned typed-ref
                // local trap per spec. funcref/externref/abstract nulls stay
                // plain (they aren't GC struct/array refs).
                s if s.contains('$') => Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_typed_null")),
                    args: vec![],
                    optional: false }),
                _ => Expression::null() };
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(var_name),
                    type_hint: Some(t.clone().into()),
                    init: Some(init),
                    array_bounds: None,
                    with_events: false }],
                kind: VarDeclKind::Let })
        })
        .collect())
}

// ── Instructions ──────────────────────────────────────────────────────────────
//
// WAT instructions that produce a control-flow effect (block, loop, br, br_if,
// return, unreachable) are lowered to proper AST *statements* so the compiler
// emits the right WASM structured-control opcodes (BLOCK, LOOP, BR, BR_IF,
// RETURN, UNREACHABLE).  Value-producing instructions become expressions.

fn walk_instr_as_stmts(
    pair: Pair<Rule>,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_stmts(inner, span, labels),
        Rule::plain_instr => walk_plain_instr_as_stmts(inner, span, labels),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())) }
}

/// The WASM trap, as the shared compiler already spells it.
///
/// A zero-argument call named after a WASM instruction is resolved straight
/// from the VM's own opcode table (`Op::from_flattened_name`) and emitted by
/// `emit_builtin_opcode` — the same route every other raw instruction in this
/// walker takes. `unreachable` is `Op::new(0x00, 0x00)`, declared in
/// `core_ops.rs`, so no new mechanism, AST node or builtin is needed: this was
/// the only front end that was not using the route it already had.
fn trap_expr() -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("unreachable")),
        args: Vec::new(),
        optional: false })
}

fn walk_instr_as_expr(pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(inner, span, labels),
        Rule::plain_instr => walk_plain_instr_as_expr(inner, span, labels),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())) }
}

// ── Plain instructions ────────────────────────────────────────────────────────

fn walk_plain_instr_as_stmts(
    pair: Pair<Rule>,
    _span: Span,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    fold_instructions(vec![pair], labels)
}

fn walk_plain_instr_as_expr(
    pair: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Expression, String> {
    let mut name = String::new();
    let mut raw_args: Vec<Pair<Rule>> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::instr_name => name = child.as_str().to_string(),
            Rule::instr_arg => raw_args.push(child),
            _ => {}
        }
    }
    // Peel any leading bare memidx immediate(s) into a `@@mem<N>` suffix first.
    let name = peel_mem_selector(&name, &mut raw_args, labels)?;
    let mut args: Vec<Expression> = Vec::new();
    for raw in raw_args {
        args.push(walk_instr_arg_pair(raw, labels)?);
    }
    map_instr_to_ast(name, args, span)
}

// ── Folded instructions ───────────────────────────────────────────────────────

fn walk_folded_instr_as_stmts(
    pair: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    // Collect all children so we can inspect the name and process the rest.
    let mut all_children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if all_children.is_empty() {
        return Ok(Vec::new());
    }

    // The name is always first if present.
    let name = if all_children[0].as_rule() == Rule::instr_name {
        all_children.remove(0).as_str().to_string()
    } else {
        String::new()
    };

    match name.as_str() {
        // ── (block $label instr*) → Labeled { label, Block([stmts]) } ─────
        "block" => {
            let mut label: Option<String> = None;
            let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
            for child in all_children {
                match child.as_rule() {
                    Rule::id => label = Some(child.as_str()[1..].to_string()),
                    Rule::block_type => {}
                    Rule::instr => instr_pairs.push(child),
                    _ => {}
                }
            }
            let effective = labels.push(label.clone(), LabelKind::Block, Vec::new());
            let body = fold_instructions(instr_pairs, labels)?;
            labels.pop();
            let block_stmt = Statement::with_span(StmtKind::Block(body), span);
            Ok(vec![Statement::with_span(
                StmtKind::Labeled {
                    label: effective,
                    body: Box::new(block_stmt) },
                span,
            )])
        }

        // ── (loop $label instr*) → Labeled { label, While(true, [stmts]) }
        "loop" => {
            let mut label: Option<String> = None;
            let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
            for child in all_children {
                match child.as_rule() {
                    Rule::id => label = Some(child.as_str()[1..].to_string()),
                    Rule::block_type => {}
                    Rule::instr => instr_pairs.push(child),
                    _ => {}
                }
            }
            let effective = labels.push(label.clone(), LabelKind::Loop, Vec::new());
            let mut body = fold_instructions(instr_pairs, labels)?;
            labels.pop();
            // A WASM loop exits on fall-through; while(true) needs an explicit break.
            body.push(Statement::with_span(
                StmtKind::Break(BreakTarget::Implicit),
                span,
            ));
            let while_stmt = Statement::with_span(
                StmtKind::While {
                    cond: Expression::bool(true),
                    body,
                    else_body: None },
                span,
            );
            Ok(vec![Statement::with_span(
                StmtKind::Labeled {
                    label: effective,
                    body: Box::new(while_stmt) },
                span,
            )])
        }

        // ── (return instr?) ───────────────────────────────────────────────
        "return" => {
            let val = all_children
                .into_iter()
                .find(|c| c.as_rule() == Rule::instr)
                .map(|c| walk_instr_as_expr(c, labels))
                .transpose()?;
            Ok(vec![Statement::with_span(StmtKind::Return(val), span)])
        }

        // ── (unreachable) = WASM trap ─────────────────────────────────────
        // A trap, NOT a throw. This was `StmtKind::Throw { expr: None }`, which
        // is an exception: `(block $d (try_table (catch_all $d) unreachable))`
        // swallowed it and the program exited 0. Per the spec a trap is outside
        // the exception system and no handler can intercept it. `Op::UNREACHABLE`
        // returns `Err` straight out of the interpreter loop and never consults
        // the handler stack.
        "unreachable" => Ok(vec![Statement::with_span(
            StmtKind::Expr(trap_expr()),
            span,
        )]),

        // ── (br $label) ───────────────────────────────────────────────────
        "br" => {
            let lbl = all_children
                .iter()
                .find(|c| c.as_rule() == Rule::instr_arg)
                .and_then(|a| a.clone().into_inner().next())
                .filter(|c| c.as_rule() == Rule::id)
                .map(|c| c.as_str()[1..].to_string());
            Ok(vec![make_br_stmt_opt(lbl.as_deref(), labels, span)])
        }

        // ── (br_if $label cond) ───────────────────────────────────────────
        "br_if" => {
            let mut lbl: Option<String> = None;
            let mut cond: Option<Expression> = None;
            for child in &all_children {
                match child.as_rule() {
                    Rule::instr_arg => {
                        if let Some(inner) = child.clone().into_inner().next() {
                            if inner.as_rule() == Rule::id && lbl.is_none() {
                                lbl = Some(inner.as_str()[1..].to_string());
                            } else if cond.is_none() {
                                cond = Some(instr_arg_inner_to_expr(inner));
                            }
                        }
                    }
                    _ => {}
                }
            }
            // condition may come from an inner instr
            if cond.is_none() {
                for child in all_children {
                    if child.as_rule() == Rule::instr {
                        cond = Some(walk_instr_as_expr(child, labels)?);
                        break;
                    }
                }
            }
            let cond_expr = cond.unwrap_or(Expression::int(0));
            let branch = make_br_stmt_opt(lbl.as_deref(), labels, span);
            Ok(vec![Statement::with_span(
                StmtKind::If {
                    cond: cond_expr,
                    then_body: vec![branch],
                    else_body: None,
                    elifs: Vec::new() },
                span,
            )])
        }

        // ── all other folded instructions → expression statement ──────────
        _ => {
            let expr = walk_folded_core(name, all_children, span, labels)?;
            Ok(vec![Statement::with_span(StmtKind::Expr(expr), span)])
        }
    }
}

fn walk_folded_instr_as_expr(
    pair: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Expression, String> {
    let pair_text = pair.as_str().to_string();
    let mut all_children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if all_children.is_empty() {
        return Ok(Expression::null());
    }
    let name = if all_children[0].as_rule() == Rule::instr_name {
        all_children.remove(0).as_str().to_string()
    } else {
        // Folded block/loop/if/try lead with a bare keyword literal (not an
        // instr_name token), so it never appears as a child — recover it from
        // the source text. Without this the head is "", and the instruction
        // falls through to an empty-callee call.
        folded_head_keyword(&pair_text).unwrap_or_default()
    };
    walk_folded_core(name, all_children, span, labels)
}

/// The head instruction of a folded_instr: its `instr_name` token, or the
/// structured keyword (`block`/`loop`/`if`/`try`) recovered from source text.
fn folded_instr_head(pair: &Pair<Rule>) -> String {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::instr_name)
        .map(|c| c.as_str().to_string())
        .or_else(|| folded_head_keyword(pair.as_str()))
        .unwrap_or_default()
}

/// The leading keyword of a folded block/loop/if/try S-expression (`(block …)`
/// → `"block"`), recovered from source text because the grammar consumes these
/// keywords as literals rather than `instr_name` tokens.
fn folded_head_keyword(text: &str) -> Option<String> {
    let rest = text.trim_start().strip_prefix('(')?.trim_start();
    // `try_table` before `if`-less list; `_` is an identifier continuation, so a
    // keyword boundary must reject it (else `try` wrongly matches `try_table`).
    ["block", "loop", "if", "try_table"].iter().find_map(|kw| {
        rest.strip_prefix(kw)
            .filter(|after| {
                after.is_empty() || after.starts_with(|c: char| !c.is_alphanumeric() && c != '_')
            })
            .map(|_| kw.to_string())
    })
}

/// Core folded instruction → expression (shared by both statement and expression contexts).
fn walk_folded_core(
    name: String,
    children: Vec<Pair<Rule>>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Expression, String> {
    // ── (block $label instr*) used as expression ──────────────────────────
    if name == "block" || name == "loop" {
        let kind = if name == "block" {
            LabelKind::Block
        } else {
            LabelKind::Loop
        };
        let mut label: Option<String> = None;
        let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
        for child in children {
            match child.as_rule() {
                Rule::id => label = Some(child.as_str()[1..].to_string()),
                Rule::instr => instr_pairs.push(child),
                _ => {}
            }
        }
        labels.push(label.clone(), kind.clone(), Vec::new());
        let body = fold_instructions(instr_pairs, labels)?;
        labels.pop();
        let last_expr = if let Some(last) = body.last() {
            if let StmtKind::Expr(e) = &last.kind {
                e.clone()
            } else {
                Expression::null()
            }
        } else {
            Expression::null()
        };
        return Ok(last_expr);
    }

    // ── (if cond (then ...) (else ...)) → ternary ─────────────────────────
    let mut args: Vec<Expression> = Vec::new();
    let mut then_exprs: Vec<Expression> = Vec::new();
    let mut else_exprs: Vec<Expression> = Vec::new();
    let mut has_then = false;

    for child in children {
        match child.as_rule() {
            Rule::instr_name => {} // already consumed
            Rule::id => {}         // label — ignore in expression context
            Rule::block_type => {} // result type annotation
            Rule::instr_arg => args.push(walk_instr_arg_pair(child, labels)?),
            Rule::instr => args.push(walk_instr_as_expr(child, labels)?),
            Rule::then_block => {
                has_then = true;
                let mut instr_pairs = Vec::new();
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        instr_pairs.push(sub);
                    }
                }
                let body = fold_instructions(instr_pairs, labels)?;
                let last_expr = if let Some(last) = body.last() {
                    if let StmtKind::Expr(e) = &last.kind {
                        e.clone()
                    } else {
                        Expression::null()
                    }
                } else {
                    Expression::null()
                };
                then_exprs.push(last_expr);
            }
            Rule::else_block => {
                let mut instr_pairs = Vec::new();
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        instr_pairs.push(sub);
                    }
                }
                let body = fold_instructions(instr_pairs, labels)?;
                let last_expr = if let Some(last) = body.last() {
                    if let StmtKind::Expr(e) = &last.kind {
                        e.clone()
                    } else {
                        Expression::null()
                    }
                } else {
                    Expression::null()
                };
                else_exprs.push(last_expr);
            }
            _ => {}
        }
    }

    if name == "if" || has_then {
        let cond = args.into_iter().next().unwrap_or(Expression::bool(false));
        let then_val = then_exprs.into_iter().last().unwrap_or(Expression::null());
        let else_val = else_exprs.into_iter().last().unwrap_or(Expression::null());
        return Ok(Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then_val),
                else_: Box::new(else_val) },
            span,
        ));
    }

    map_instr_to_ast(name, args, span)
}

/// Build an expression that renders `x` the way the WAT text format prints a
/// float: NaN → "nan", ±∞ → "inf"/"-inf", whole numbers gain a ".0" suffix, and
/// everything else uses the natural decimal. Uses only native operators (so it
/// needs the profile's `dynamic_add` for the string concatenation) — no host
/// helpers. `x` is pure arithmetic at every call site, so re-reading it is safe.
fn wat_float_format(x: Expression) -> Expression {
    fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r) })
    }
    fn tern(c: Expression, t: Expression, e: Expression) -> Expression {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(c),
            then: Box::new(t),
            else_: Box::new(e) })
    }
    let zero = || Expression::float(0.0);
    // whole number → "<x>.0"; otherwise the natural decimal string.
    let finite = tern(
        bin(
            BinOp::StrictEq,
            bin(BinOp::Mod, x.clone(), Expression::float(1.0)),
            zero(),
        ),
        bin(
            BinOp::Add,
            bin(BinOp::Add, x.clone(), Expression::string("")),
            Expression::string(".0"),
        ),
        bin(BinOp::Add, x.clone(), Expression::string("")),
    );
    // (x - x) is 0 for finite values but NaN for ±∞.
    let inf_or_finite = tern(
        bin(
            BinOp::StrictNotEq,
            bin(BinOp::Sub, x.clone(), x.clone()),
            zero(),
        ),
        tern(
            bin(BinOp::Lt, x.clone(), zero()),
            Expression::string("-inf"),
            Expression::string("inf"),
        ),
        finite,
    );
    // NaN is the only value not equal to itself.
    tern(
        bin(BinOp::StrictNotEq, x.clone(), x.clone()),
        Expression::string("nan"),
        inf_or_finite,
    )
}

// ── map_instr_to_ast — WAT instruction name → common AST expression ───────────

/// Strip `offset=`/`align=` memarg string-args from a load/store's operand list
/// and fold a non-zero `offset=N` into the address (the first stack operand):
/// `i32.load offset=5` over base `a` becomes a load of `a + 5`.
fn fold_memarg_offset(args: Vec<Expression>, span: Span) -> Vec<Expression> {
    let mut offset: i64 = 0;
    let mut rest: Vec<Expression> = Vec::new();
    for a in args {
        if let ExprKind::Lit(Literal::Str(s)) = &a.kind {
            if let Some(n) = s.strip_prefix("offset=") {
                offset += n.parse::<i64>().unwrap_or(0);
                continue;
            }
            if s.starts_with("align=") {
                continue;
            }
        }
        rest.push(a);
    }
    if offset != 0 {
        if let Some(addr) = rest.first().cloned() {
            rest[0] = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(addr),
                    right: Box::new(Expression::int(offset)) },
                span,
            );
        }
    }
    rest
}

/// The binding name a numeric `local.get`/`global.get` (and set) index lowers
/// to. Locals/params use the `p<i>` name space; globals resolve through the
/// declaration-order `GLOBAL_INDEX_NAME` (falling back to the same synthetic
/// scheme the pre-scan used, so it works even if the pre-scan missed one).
fn index_binding_name(i: i64, is_global: bool) -> String {
    if is_global {
        GLOBAL_INDEX_NAME
            .with(|g| g.borrow().get(i as usize).cloned())
            .unwrap_or_else(|| format!("__wasm_global_{i}"))
    } else {
        format!("p{i}")
    }
}

/// Resolve a memory-index immediate: a literal integer, or a `$name` looked up
/// in the declaration-order `MEMORY_NAME_INDEX`. Anything else defaults to 0.
fn resolve_wat_memidx(e: &Expression) -> usize {
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => *n as usize,
        ExprKind::Ident(nm) => MEMORY_NAME_INDEX
            .with(|f| f.borrow().get(nm).copied())
            .unwrap_or(0),
        _ => 0 }
}

/// Number of leading memory-index immediates a memory op may carry: `memory.copy`
/// names two memories (dst, src); every other memory op names at most one. 0 =
/// not a memory op (never peels a selector).
fn mem_op_immediate_count(name: &str) -> usize {
    match name {
        "memory.copy" => 2,
        "memory.fill" | "memory.size" | "memory.grow" | "i32.load" | "i64.load" | "f32.load"
        | "f64.load" | "i32.load8_s" | "i32.load8_u" | "i32.load16_s" | "i32.load16_u"
        | "i64.load8_s" | "i64.load8_u" | "i64.load16_s" | "i64.load16_u" | "i64.load32_s"
        | "i64.load32_u" | "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8"
        | "i32.store16" | "i64.store8" | "i64.store16" | "i64.store32" => 1,
        _ => 0 }
}

/// True when a raw `instr_arg` is a BARE index immediate (`integer`/`id`) — i.e.
/// a memory index — as opposed to a folded operand (`folded_instr`) or an
/// `offset=`/`align=` memarg. This is the reliable signal distinguishing a real
/// memidx from an operand the WAT grammar greedily attaches as an `instr_arg`.
fn is_bare_index_arg(raw: &Pair<Rule>) -> bool {
    matches!(
        raw.clone().into_inner().next().map(|x| x.as_rule()),
        Some(Rule::integer) | Some(Rule::id)
    )
}

/// Peel leading bare memory-index immediates off a memory op's raw `instr_args`,
/// returning the `@@mem<N>`-mangled op name (unchanged when only the default
/// memory is named). `raw_args` is left holding just the real operand args, so
/// no operand is ever mistaken for a selector. The compiler turns each `@@mem<N>`
/// into the VM's fixed 4-byte selector.
fn peel_mem_selector(
    name: &str,
    raw_args: &mut Vec<Pair<Rule>>,
    labels: &mut LabelStack,
) -> Result<String, String> {
    let n = mem_op_immediate_count(name);
    if n == 0 {
        return Ok(name.to_string());
    }
    let mut indices = Vec::new();
    while indices.len() < n && raw_args.first().map(is_bare_index_arg).unwrap_or(false) {
        let r = raw_args.remove(0);
        let e = walk_instr_arg_pair(r, labels)?;
        indices.push(resolve_wat_memidx(&e));
    }
    if indices.iter().all(|&i| i == 0) {
        // Only the default memory (or none) named — the bare immediates were
        // still consumed above (they are selectors, not operands).
        return Ok(name.to_string());
    }
    if name == "memory.copy" {
        let dst = indices.first().copied().unwrap_or(0);
        let src = indices.get(1).copied().unwrap_or(0);
        Ok(format!("memory.copy@@mem{dst}@@mem{src}"))
    } else {
        Ok(format!("{}@@mem{}", name, indices[0]))
    }
}

fn map_instr_to_ast(name: String, args: Vec<Expression>, span: Span) -> Result<Expression, String> {
    // A constant `offset=N` memarg on a load/store folds into the address
    // (WASM effective address = base + offset). The VM's load/store opcode
    // stream can't unambiguously carry a memarg — its optional-memarg peek
    // can't tell offset bytes from the next opcode — and a static offset needs
    // no runtime immediate. `align=` memargs are pure hints and are dropped.
    let args = if name.contains(".load") || name.contains(".store") {
        fold_memarg_offset(args, span)
    } else {
        args
    };
    // NOTE: multi-memory selectors (`i32.store 1`, `memory.copy 1 0`, …) are
    // peeled off at the plain-instruction parse site (`peel_mem_selector`), where
    // a genuine memidx immediate — a BARE `integer`/`id` token — is distinguishable
    // from a folded operand the WAT grammar greedily attaches as an `instr_arg`.
    // By the time args reach here they are already `name@@mem<N>`-mangled with the
    // memidx stripped, so no arity-based inference (which cannot tell an immediate
    // from a flushed operand) is done in this shared lowering.
    match name.as_str() {
        // `table.init tableidx? elemidx` — WAT allows 1 index (elemidx, table 0)
        // or 2 (tableidx elemidx). The VM's TABLE_INIT reads two byte immediates
        // in [elem_idx, table_idx] order, then 3 stack operands (dst, src, len).
        // Normalize the leading immediates to exactly [elem, table] so the table
        // index is never mistaken for the first stack operand.
        "table.init" => {
            let mut a = args;
            let n_idx = a.len().saturating_sub(3); // 3 stack operands
            let (elem, table) = match n_idx {
                0 => (Expression::int(0), Expression::int(0)),
                1 => (a.remove(0), Expression::int(0)),
                _ => {
                    let table = a.remove(0); // text order: tableidx first
                    let elem = a.remove(0); // then elemidx
                    (elem, table)
                }
            };
            let mut new_args = vec![elem, table];
            new_args.append(&mut a);
            return Ok(make_call("table.init", new_args, span));
        }
        // Typeless array access: the WAT typeidx (`$t`) immediates are the first
        // arg(s) but the VM's array.get/set/fill/copy don't read them — drop and
        // keep only the stack operands. array.copy carries two typeidxs.
        // `array.get`/`array.set`/`array.fill`: the WAT typeidx is dropped; the
        // compiler's `emit_named_opcode` traps on null/out-of-bounds for a spec
        // (`function_references`) profile — see `array_get`/`array_set` there.
        "array.get" | "array.set" | "array.fill" => {
            let rest: Vec<Expression> = args.into_iter().skip(1).collect();
            Ok(make_call(&name.replace('.', "_"), rest, span))
        }
        // Packed-array reads: `array.get_s`/`array.get_u $T` read a packed `i8`/
        // `i16` element and sign-/zero-extend it to i32. The VM stores the array
        // untyped, so the width comes from the `$T` element type — plain
        // `array_get` then an extend (signed) or a mask (unsigned).
        "array.get_s" | "array.get_u" => {
            let signed = name == "array.get_s";
            let elem = args.first().and_then(|a| match &a.kind {
                ExprKind::Ident(n) => ARRAY_ELEM_TYPE.with(|m| m.borrow().get(n).cloned()),
                _ => None });
            let rest: Vec<Expression> = args.into_iter().skip(1).collect();
            let get = make_call("array_get", rest, span);
            Ok(match (elem.as_deref(), signed) {
                (Some("i8"), true) => make_call("i32.extend8_s", vec![get], span),
                (Some("i16"), true) => make_call("i32.extend16_s", vec![get], span),
                (Some("i8"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(get),
                    right: Box::new(Expression::int(0xFF)) }),
                (Some("i16"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(get),
                    right: Box::new(Expression::int(0xFFFF)) }),
                // Non-packed (i32/i64/ref/…): the stored value is already exact.
                _ => get })
        }
        "array.copy" => {
            let rest: Vec<Expression> = args.into_iter().skip(2).collect();
            Ok(make_call("array_copy", rest, span))
        }
        // ── Constants ─────────────────────────────────────────────────────
        // i32.const carries a 32-bit pattern: reinterpret the (possibly
        // unsigned, e.g. 0x80000000) literal into signed i32 range so it stays
        // exactly representable and the i32 opcodes read the right bits.
        "i32.const" => {
            let v = args.into_iter().next().unwrap_or(Expression::int(0));
            if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                let reinterp = (*n as u32) as i32 as i64;
                Ok(Expression::with_span(
                    ExprKind::Lit(Literal::Int(reinterp)),
                    span,
                ))
            } else {
                Ok(v)
            }
        }
        // i64.const needs an exact 64-bit value; a plain Int literal compiles to
        // f64 (losing bits past 2^53), so carry it as the exact-integer literal
        // the i64 opcodes read via `as_i64`.
        "i64.const" => {
            let v = args.into_iter().next().unwrap_or(Expression::int(0));
            if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                Ok(Expression::with_span(
                    ExprKind::Lit(Literal::BigInt(*n)),
                    span,
                ))
            } else {
                Ok(v)
            }
        }
        "f64.const" => Ok(args.into_iter().next().unwrap_or(Expression::float(0.0))),
        // f32.const carries an f32 value: demote the (exact-text) f64 literal to
        // single precision so it lands as `Value::F32`, matching WASM.
        "f32.const" => {
            let v = args.into_iter().next().unwrap_or(Expression::float(0.0));
            Ok(make_call("f32_demote_f64", vec![v], span))
        }
        // wasm:js-string builtins — string.const "text" → string literal
        "string.const" => Ok(args.into_iter().next().unwrap_or(Expression::string(""))),

        // ── Local / global get → Ident ────────────────────────────────────
        // A numeric index names a LOCAL/param (`p<i>`) for `local.get`, but a
        // GLOBAL by declaration index for `global.get` (separate name spaces).
        "local.get" | "global.get" => {
            let is_global = name == "global.get";
            let idx = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(match &idx.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(index_binding_name(*i, is_global)), span)
                }
                _ => idx })
        }

        // ── Local / global set → Assign ───────────────────────────────────
        "local.set" | "global.set" => {
            let is_global = name == "global.set";
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target = match &target_raw.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(index_binding_name(*i, is_global)), span)
                }
                _ => target_raw };
            Ok(Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(target),
                    value: Box::new(value) },
                span,
            ))
        }

        // ── local.tee → assign + value ────────────────────────────────────
        "local.tee" => {
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target_name = match &target_raw.kind {
                ExprKind::Ident(n) => n.clone(),
                ExprKind::Lit(Literal::Int(i)) => format!("p{}", i),
                _ => "__tee_tmp".to_string() };
            Ok(Expression::with_span(
                ExprKind::Sequence(vec![
                    Expression::with_span(
                        ExprKind::Assign {
                            target: Box::new(Expression::ident(&target_name)),
                            value: Box::new(value) },
                        span,
                    ),
                    Expression::ident(&target_name),
                ]),
                span,
            ))
        }

        // ── Binary arithmetic ─────────────────────────────────────────────
        // Every typed WASM op routes to its real opcode (via the default
        // make_call below → profile `opcode:<op>`) so the VM applies genuine
        // WASM semantics — i32/i64 wrapping and signed/unsigned splits, f32
        // single precision. Only f64 arithmetic (native IEEE double) and float
        // comparisons stay on the shared BinOp path, where they are exact.
        "f64.add" => bin_op(args, BinOp::Add, span),
        "f64.sub" => bin_op(args, BinOp::Sub, span),
        "f64.mul" => bin_op(args, BinOp::Mul, span),
        "f64.div" => bin_op(args, BinOp::Div, span),

        // ── Comparisons ───────────────────────────────────────────────────
        // Float comparisons compare exact f64 widenings; i32/i64 comparisons
        // (with signed/unsigned variants) route to their opcodes below.
        "f32.eq" | "f64.eq" => bin_op(args, BinOp::Eq, span),
        "f32.ne" | "f64.ne" => bin_op(args, BinOp::NotEq, span),
        "f32.lt" | "f64.lt" => bin_op(args, BinOp::Lt, span),
        "f32.gt" | "f64.gt" => bin_op(args, BinOp::Gt, span),
        "f32.le" | "f64.le" => bin_op(args, BinOp::LtEq, span),
        "f32.ge" | "f64.ge" => bin_op(args, BinOp::GtEq, span),

        // i32.eqz / i64.eqz route to their opcodes (default make_call below).

        // ── Unary negation ────────────────────────────────────────────────
        // f32.neg routes to the f32 opcode (default make_call below) so it
        // yields a single-precision Value::F32; f64.neg uses the AST Unary.
        "f64.neg" => {
            let operand = args.into_iter().next().unwrap_or(Expression::float(0.0));
            Ok(Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(operand) },
                span,
            ))
        }

        // ── select → ternary ──────────────────────────────────────────────
        "select" => {
            // `select (result t)` prepends a result-type annotation; the stack
            // operands (val1, val2, cond) are always the last three args.
            let n = args.len();
            let val1 = args
                .get(n.wrapping_sub(3))
                .cloned()
                .unwrap_or(Expression::null());
            let val2 = args
                .get(n.wrapping_sub(2))
                .cloned()
                .unwrap_or(Expression::null());
            let cond = args
                .get(n.wrapping_sub(1))
                .cloned()
                .unwrap_or(Expression::bool(false));
            Ok(Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(cond),
                    then: Box::new(val1),
                    else_: Box::new(val2) },
                span,
            ))
        }

        // ── drop → evaluate and discard ───────────────────────────────────
        "drop" => Ok(args.into_iter().next().unwrap_or(Expression::null())),

        // ── nop ───────────────────────────────────────────────────────────
        "nop" => Ok(Expression::with_span(ExprKind::Lit(Literal::Null), span)),

        // ── unreachable in expression context ─────────────────────────────
        // A folded `(unreachable)` used as a VALUE used to compile to
        // `Expression::null()` — nothing at all. `(func $f (result i32)
        // (unreachable))` returned null and the caller carried on to exit 0, so
        // any wast test whose failure path was written that way passed
        // unconditionally. It traps here exactly as in statement position.
        "unreachable" => Ok(trap_expr()),

        // ── return / br in expression context ─────────────────────────────
        // These are meaningful at statement level; here they produce null.
        "return" | "br" | "br_if" | "br_table" => Ok(Expression::null()),

        // ── call → Call(callee, args) ─────────────────────────────────────
        "call" => {
            let mut it = args.into_iter();
            let callee = it.next().unwrap_or(Expression::null());
            let mut call_args: Vec<Expression> = it.collect();
            // `$log_f64` prints in WAT text (`4.0`, `inf`, `nan`) rather than
            // ECMA `ToString` (`4`, `Infinity`); pre-format its f64 argument.
            // `$log_f32` needs no wrapper — a `Value::F32` already Displays as
            // WAT float text (f32 is a WASM-only value type).
            if let ExprKind::Ident(n) = &callee.kind {
                if n == "log_f64" {
                    call_args = call_args.into_iter().map(wat_float_format).collect();
                }
            }
            // A call to a function DEFINED in this module targets a static method
            // of the module class; qualify `Ident(f)` as `ClassName.f`. Imports
            // keep their bare name so the profile builtin table resolves them.
            let callee = match &callee.kind {
                ExprKind::Ident(n) if DEFINED_FUNC_NAMES.with(|d| d.borrow().contains(n)) => {
                    let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(Expression::ident(&class)),
                            field: n.clone(),
                            null_safe: false },
                        span,
                    )
                }
                _ => callee };
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: call_args.into_iter().map(Argument::positional).collect(),
                    optional: false },
                span,
            ))
        }

        // ── GC / WasmGC reference ops ─────────────────────────────────────
        // ref.null <heaptype> pushes a typed null reference. The heap type is
        // an immediate annotation, not a stack value, and the VM has a single
        // null — so drop the arg and produce a plain null (like `nop`). Applies
        // to bare heap types (`func`/`extern`) and indexed types (`$T`) alike.
        // `ref.null $t` → a WASM GC typed null (traps on struct.get/array.get
        // per spec), distinct from a plain null. Lowered to the compiler builtin
        // `__wast_typed_null` which emits `ref.null` (Op::NULL) with a non-zero
        // heap-type immediate so the VM produces a `TypedNull`.
        "ref.null" => Ok(make_call("__wast_typed_null", vec![], span)),
        // ref.func $f → a first-class reference to module function `$f`. Module
        // functions are static methods of the module class, so this is the
        // static method referenced as a value (the compiler tears it off into a
        // funcref). ref.func by numeric index is not resolved here (needs the
        // compiler's chunk table); named refs cover the common case.
        "ref.func" => {
            let field = match args.into_iter().next() {
                Some(e) => match &e.kind {
                    ExprKind::Ident(n) => n.clone(),
                    _ => return Ok(e) },
                None => return Ok(Expression::null()) };
            let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
            Ok(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::ident(&class)),
                    field,
                    null_safe: false },
                span,
            ))
        }
        // ref.extern N: a WAST-harness host externref carrying the integer
        // payload N (used to create/compare externref values in assert scripts).
        // The VM has no host-externref type; model it faithfully as its integer
        // payload so equality-by-payload works both as an invoke arg and as an
        // expected result. `result_val`'s `(ref.extern N)` already lowers to N.
        "ref.extern" => Ok(args
            .into_iter()
            .next()
            .unwrap_or_else(|| Expression::int(0))),
        // call_ref $sig: call a funcref value. args = [$sig, ...operands]; the
        // funcref is on top of the stack (last operand), the sig's params
        // precede it. Lower to a Call on the funcref value (compiler → CALL_REF).
        "call_ref" => {
            let mut rest = args;
            if !rest.is_empty() {
                rest.remove(0); // drop the $sig type immediate
            }
            let callee = rest.pop().unwrap_or_else(Expression::null);
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: rest.into_iter().map(Argument::positional).collect(),
                    optional: false },
                span,
            ))
        }

        // ── GC / WasmGC struct ops ────────────────────────────────────────
        // struct.new $T v0 v1 ...  → {"0": v0, "1": v1, ..., "__type": "T"}
        // args: [typeidx, field_val_0, field_val_1, ...]. The `__type` stamp
        // carries the GC type name so the VM's `ref.test`/`ref.cast`/`br_on_cast`
        // resolve identity + subtyping through the registered type hierarchy.
        "struct.new" => {
            let type_name = args.first().map(wasm_type_ref_name).unwrap_or_default();
            let vals: Vec<Expression> = if args.len() > 1 {
                args[1..].to_vec()
            } else {
                vec![]
            };
            let props: Vec<ObjectProperty> = vals
                .into_iter()
                .enumerate()
                .map(|(i, v)| ObjectProperty::KeyValue {
                    key: Expression::string(&i.to_string()),
                    value: v })
                .collect();
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_type(obj, &type_name, span))
        }
        // struct.new_default $T → each field set to its storage type's default
        // (0 for ints, 0.0 for floats, null for refs), stamped with its rtt.
        "struct.new_default" => {
            let type_name = args.first().map(wasm_type_ref_name).unwrap_or_default();
            let field_types = STRUCT_FIELD_TYPES
                .with(|m| m.borrow().get(&type_name).cloned())
                .unwrap_or_default();
            let props: Vec<ObjectProperty> = field_types
                .iter()
                .enumerate()
                .map(|(i, ty)| ObjectProperty::KeyValue {
                    key: Expression::string(&i.to_string()),
                    value: default_value_for_storage_type(ty) })
                .collect();
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_type(obj, &type_name, span))
        }
        // array.new_default $T → a length-N array filled with the element type's
        // default. For numeric elements that's 0/0.0, so lower to `array.new $T
        // <default> <length>` (which fills the value); ref elements keep the VM's
        // null-fill via `array_new_default`. args: [typeidx, length].
        "array.new_default" => {
            let elem = args.first().and_then(|a| match &a.kind {
                ExprKind::Ident(n) => ARRAY_ELEM_TYPE.with(|m| m.borrow().get(n).cloned()),
                _ => None });
            // Numeric OR concrete `(ref null $t)` elements have a known default
            // (0/0.0/typed-null), so lower to `array.new $T <default> <length>`
            // which fills every lane. funcref/externref/unknown keep the VM's
            // plain-null fill via `array_new_default`.
            let has_known_default = matches!(
                elem.as_deref(),
                Some("i8" | "i16" | "i32" | "i64" | "f32" | "f64")
            ) || elem.as_deref().is_some_and(|s| s.contains('$'));
            if has_known_default {
                let typeidx = args.first().cloned().unwrap_or(Expression::int(0));
                let length = args.into_iter().nth(1).unwrap_or(Expression::int(0));
                let default = default_value_for_storage_type(elem.as_deref().unwrap_or(""));
                Ok(make_call("array_new", vec![typeidx, default, length], span))
            } else {
                Ok(make_call("array_new_default", args, span))
            }
        }
        // array.new_fixed $T N v0 v1 … → [v0, v1, …] stamped with $T's rtt so
        // `array.get`/`set` trap on OOB (WASM GC). args: [typeidx, N, v0…]; the
        // N stack values become an array literal, then `__wast_stamp_array_type`
        // registers $T and stamps its type id.
        "array.new_fixed" => {
            let type_name = args.first().map(wasm_type_ref_name).unwrap_or_default();
            let vals: Vec<ArrayElement> = if args.len() > 2 {
                args[2..]
                    .iter()
                    .map(|v| ArrayElement {
                        key: None,
                        value: v.clone(),
                        spread: false,
                        by_ref: false })
                    .collect()
            } else {
                vec![]
            };
            let arr = Expression::with_span(ExprKind::Array(vals), span);
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_stamp_array_type")),
                    args: vec![
                        Argument::positional(arr),
                        Argument::positional(Expression::string(&type_name)),
                    ],
                    optional: false },
                span,
            ))
        }
        // struct.get $T N ref  → ref["N"] (null-trapped). The `_s`/`_u` variants
        // sign/zero-extend a packed i8/i16 field, mirroring array.get_s/get_u.
        // args: [typeidx, fieldidx, ref_expr]
        "struct.get" | "struct.get_s" | "struct.get_u" => {
            let type_name = args.first().map(wasm_type_ref_name).unwrap_or_default();
            let field_idx = args
                .get(1)
                .and_then(|a| {
                    if let ExprKind::Lit(Literal::Int(i)) = &a.kind {
                        Some(*i)
                    } else {
                        None
                    }
                })
                .unwrap_or(0);
            let obj = args.into_iter().nth(2).unwrap_or(Expression::null());
            let member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(obj),
                    field: field_idx.to_string(),
                    null_safe: false },
                span,
            );
            if name == "struct.get" {
                return Ok(member);
            }
            let signed = name == "struct.get_s";
            let field_ty = STRUCT_FIELD_TYPES.with(|m| {
                m.borrow()
                    .get(&type_name)
                    .and_then(|v| v.get(field_idx as usize).cloned())
            });
            Ok(match (field_ty.as_deref(), signed) {
                (Some("i8"), true) => make_call("i32.extend8_s", vec![member], span),
                (Some("i16"), true) => make_call("i32.extend16_s", vec![member], span),
                (Some("i8"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(member),
                    right: Box::new(Expression::int(0xFF)) }),
                (Some("i16"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(member),
                    right: Box::new(Expression::int(0xFFFF)) }),
                // Non-packed (i32/i64/ref/…): the stored value is already exact.
                _ => member })
        }
        // struct.set $T N ref val → ref["N"] = val  (produces null, used as stmt)
        // args: [typeidx, fieldidx, ref_expr, val_expr]
        "struct.set" => {
            let field_idx = args
                .get(1)
                .and_then(|a| {
                    if let ExprKind::Lit(Literal::Int(i)) = &a.kind {
                        Some(*i)
                    } else {
                        None
                    }
                })
                .unwrap_or(0);
            let obj = args.get(2).cloned().unwrap_or(Expression::null());
            let val = args.into_iter().nth(3).unwrap_or(Expression::null());
            Ok(Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(obj),
                            field: field_idx.to_string(),
                            null_safe: false },
                        span,
                    )),
                    value: Box::new(val) },
                span,
            ))
        }

        // ── everything else → call with dots replaced by underscores ──────
        _ => Ok(make_call(&name.replace('.', "_"), args, span)) }
}

// ── Instruction argument helpers ──────────────────────────────────────────────

fn walk_instr_arg_pair(pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty instr_arg")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(inner, Span::default(), labels),
        _ => Ok(instr_arg_inner_to_expr(inner)) }
}

fn instr_arg_inner_to_expr(inner: Pair<Rule>) -> Expression {
    match inner.as_rule() {
        Rule::float => parse_float(inner.as_str()),
        Rule::integer => parse_integer(inner.as_str()),
        Rule::string => Expression::string(&unquote(inner.as_str())),
        Rule::id => Expression::ident(&inner.as_str()[1..]),
        Rule::val_type
        | Rule::bare_val_type
        | Rule::bare_lane_type
        | Rule::bare_heap_type
        | Rule::mem_arg
        | Rule::val_lane_type => Expression::string(inner.as_str()),
        _ => Expression::null() }
}

// ── Break/continue helper ─────────────────────────────────────────────────────

fn make_br_stmt_opt(label: Option<&str>, labels: &LabelStack, span: Span) -> Statement {
    match label {
        Some(lbl) => match labels.kind_of(lbl) {
            Some(LabelKind::Loop) => Statement::with_span(
                StmtKind::Continue(ContinueTarget::Label(lbl.to_string())),
                span,
            ),
            _ => Statement::with_span(StmtKind::Break(BreakTarget::Label(lbl.to_string())), span) },
        None => Statement::with_span(StmtKind::Break(BreakTarget::Implicit), span) }
}

// ── Module fields ─────────────────────────────────────────────────────────────

fn walk_import_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut module_str = String::new();
    let mut name_str = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            let s = unquote(child.as_str());
            if module_str.is_empty() {
                module_str = s;
            } else if name_str.is_empty() {
                name_str = s;
            }
        }
    }
    Ok(Statement::new(StmtKind::Expr(make_call(
        "__wasm_import",
        vec![
            Expression::string(&module_str),
            Expression::string(&name_str),
        ],
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
    Ok(make_call(
        "__wasm_export",
        vec![Expression::string(&export_name)],
        Span::default(),
    ))
}

/// The binding name a global lowers to: its `$id`, or a declaration-index-unique
/// synthetic `__wasm_global_<idx>` when unnamed (so multiple unnamed globals do
/// not collide and `global.get <idx>` can resolve to it).
fn global_binding_name(pair: &Pair<Rule>, idx: usize) -> String {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
        .unwrap_or_else(|| format!("__wasm_global_{idx}"))
}

fn walk_global_field(pair: Pair<Rule>, idx: usize) -> Result<(String, Expression), String> {
    let name = global_binding_name(&pair, idx);
    let mut init = Expression::int(0);
    let mut labels = LabelStack::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::instr {
            init = walk_instr_as_expr(child, &mut labels)?;
        }
    }
    Ok((name, init))
}

// ── Linear memory + data segments ─────────────────────────────────────────────

fn walk_memory_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut min_pages: u64 = 0;
    let mut max_pages: Option<u64> = None;
    for child in pair.into_inner() {
        if child.as_rule() == Rule::mem_type {
            let mut nums = child.into_inner().filter(|p| p.as_rule() == Rule::integer);
            if let Some(min) = nums.next() {
                min_pages = parse_wat_u64(min.as_str());
            }
            if let Some(max) = nums.next() {
                max_pages = Some(parse_wat_u64(max.as_str()));
            }
        }
    }
    Ok(Statement::with_span(
        StmtKind::MemoryDecl {
            min_pages,
            max_pages },
        span,
    ))
}

/// Extract `(tag $e (param t*))`'s name (without `$`) and payload arity.
fn scan_tag_signature(pair: Pair<Rule>) -> (Option<String>, u8) {
    let mut name: Option<String> = None;
    let mut arity: u8 = 0;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::tag_type => {
                // tag_type = ("func" param*) | param* — count val types across params.
                arity = child
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::param)
                    .map(|p| {
                        p.into_inner()
                            .filter(|v| v.as_rule() == Rule::any_val_type)
                            .count()
                    })
                    .sum::<usize>() as u8;
            }
            _ => {}
        }
    }
    (name, arity)
}

/// `(tag $e (param t*))` — an exception-tag declaration. Emits a `WasmTagDecl`
/// the compiler imports as a load-time tag entity. Arities are recorded in the
/// module pre-scan (so `catch $e` sees them regardless of source order).
fn walk_tag_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let (name, arity) = scan_tag_signature(pair);
    Ok(Statement::with_span(
        StmtKind::WasmTagDecl {
            name: name.unwrap_or_default(),
            arity },
        span,
    ))
}

/// Walk a `(table …)` field. Returns the table declaration (goes BEFORE the
/// module class, since the VM allocates the table at instantiation) and, for the
/// inline `(table t (elem $f …))` abbreviation, its active-segment population
/// (goes AFTER the class — it references the funcs as static methods, so it must
/// run once the class exists, exactly like a standalone `(elem …)` field).
fn walk_table_field(
    pair: Pair<Rule>,
    table_idx: usize,
) -> Result<(Statement, Vec<Statement>), String> {
    let span = to_span(&pair);
    let mut min_size: u64 = 0;
    let mut max_size: Option<u64> = None;
    let mut has_table_type = false;
    // Inline `(table t (elem $f …))` abbreviation: the `index*` funcidx list.
    let mut inline_funcs: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::table_type => {
                // table_type = integer integer? ref_type — min then optional max.
                has_table_type = true;
                let mut nums = child.into_inner().filter(|p| p.as_rule() == Rule::integer);
                if let Some(min) = nums.next() {
                    min_size = parse_wat_u64(min.as_str());
                }
                if let Some(max) = nums.next() {
                    max_size = Some(parse_wat_u64(max.as_str()));
                }
            }
            Rule::index => inline_funcs.push(child.as_str().trim_start_matches('$').to_string()),
            _ => {}
        }
    }

    // Inline elem abbreviation ≡ a table sized to the element count plus an
    // active elem segment populating it from slot 0.
    if !has_table_type && !inline_funcs.is_empty() {
        let n = inline_funcs.len() as u64;
        let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
        let decl = Statement::with_span(
            StmtKind::TableDecl {
                min_size: n,
                max_size: Some(n) },
            span,
        );
        let mut population = Vec::new();
        for (i, f) in inline_funcs.iter().enumerate() {
            let funcref = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class)),
                field: f.clone(),
                null_safe: false });
            let call = make_call(
                "table_set",
                vec![
                    Expression::int(table_idx as i64),
                    Expression::int(i as i64),
                    funcref,
                ],
                span,
            );
            population.push(Statement::new(StmtKind::Expr(call)));
        }
        return Ok((decl, population));
    }

    Ok((
        Statement::with_span(StmtKind::TableDecl { min_size, max_size }, span),
        Vec::new(),
    ))
}

/// The first integer literal within `pair`'s descendants — used to read a
/// segment's constant offset (`(i32.const N)`).
fn find_first_integer(pair: &Pair<Rule>) -> Option<i64> {
    for c in pair.clone().into_inner() {
        if c.as_rule() == Rule::integer {
            if let Ok(v) = c.as_str().parse::<i64>() {
                return Some(v);
            }
        }
        if let Some(v) = find_first_integer(&c) {
            return Some(v);
        }
    }
    None
}

/// First `$id` or numeric funcidx found anywhere under `pair`, without its `$`.
/// Used to pull the funcidx out of an element initializer (`(ref.func $f)`).
/// Textual form of a WASM type-reference operand — a symbolic id (`$Sub`
/// arrives as `Ident("Sub")`) or a numeric type index. Used to name the GC
/// struct type a `struct.new`/`ref.test`/`ref.cast` refers to.
fn wasm_type_ref_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(n) => n.clone(),
        ExprKind::Lit(Literal::Int(i)) => i.to_string(),
        _ => String::new() }
}

/// Stamp a freshly-built struct object with its WASM GC rtt (the registered
/// type's id): `__wast_stamp_type(obj, "T")` → the compiler emits
/// `GLOBAL_GET __tid_T` + `SET_TYPE_ID`, so the instance carries the real
/// `type_id` the VM's `ref.test`/`ref.cast`/`is_subtype` read — no `__type`
/// string. This is the struct analogue of `array.new`'s rtt stamp.
fn wast_stamp_type(obj: Expression, type_name: &str, span: Span) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident("__wast_stamp_type")),
            args: vec![
                Argument::positional(obj),
                Argument::positional(Expression::string(type_name)),
            ],
            optional: false },
        span,
    )
}

fn first_ident_or_index(pair: &Pair<Rule>) -> Option<String> {
    for c in pair.clone().into_inner() {
        if matches!(c.as_rule(), Rule::id | Rule::index | Rule::integer) {
            let s = c.as_str().trim_start_matches('$');
            if !s.is_empty() {
                return Some(s.to_string());
            }
        }
        if let Some(v) = first_ident_or_index(&c) {
            return Some(v);
        }
    }
    None
}

/// `(elem (i32.const N) $f0 $f1 …)` — an active element segment initialising a
/// funcref table. Lowered to load-time `table.set(N+i, ref.func $fi)` for each
/// entry: the `ref.func` tear-off (Member value → REF_FUNC) produces a real
/// funcref, and `table.set` stores it, so `call_indirect` finds it at runtime.
/// (Default table 0; explicit `(table $t)` targets are TODO.)
fn walk_elem_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut offset: i64 = 0;
    let mut table_index: i64 = 0;
    let mut funcs: Vec<String> = Vec::new();
    // Only ACTIVE segments (those with an offset / `(table …)(offset …)` mode)
    // populate a table at load time. A `declare` segment merely permits
    // `ref.func` (and usually declares no table); a passive segment is copied
    // later by an explicit `table.init`. Neither should emit `table.set`.
    let mut is_active = false;
    let mut is_declare = false;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::elem_mode => {
                if child.as_str().trim().starts_with("declare") {
                    // declarative — no table population
                    is_declare = true;
                } else {
                    offset = find_first_integer(&child).unwrap_or(0);
                    // `(table $t)(offset …)` targets a NAMED table; resolve it
                    // to its declaration index (default table 0 otherwise).
                    if let Some(tname) = child
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::index)
                        .map(|i| i.as_str().trim_start_matches('$').to_string())
                    {
                        table_index = resolve_table_index(&tname);
                    }
                    is_active = true;
                }
            }
            Rule::index => funcs.push(child.as_str().trim_start_matches('$').to_string()),
            // `(item (ref.func $f))` or a bare `(ref.func $f)` element: the
            // initializer is a `ref.func` whose funcidx is the first id/index
            // inside. (Other const-expr elements have no funcref to bind.)
            Rule::elem_item => {
                if let Some(f) = first_ident_or_index(&child) {
                    funcs.push(f);
                }
            }
            _ => {}
        }
    }
    // Every element segment (active / passive / declarative) occupies one slot
    // in the element index space, in declaration order.
    let seg_index = ELEM_SEG_COUNTER.with(|c| {
        let i = *c.borrow();
        *c.borrow_mut() = i + 1;
        i
    });
    if !is_active {
        if is_declare {
            // Declarative: only permits `ref.func`, no runtime payload.
            return Ok(Statement::with_span(StmtKind::Block(Vec::new()), span));
        }
        // Passive: register the funcref list under this segment index so a later
        // `table.init $e` / `array.new_elem $e` copies real funcrefs from it.
        // Compile-time directive resolved to function chunk indices; the VM
        // materializes the funcrefs at instantiation (see `passive_elem_funcs`).
        let mut args = vec![Expression::int(seg_index as i64)];
        for f in &funcs {
            args.push(Expression::string(f));
        }
        return Ok(Statement::with_span(
            StmtKind::Expr(make_call("__wast_register_passive_elem", args, span)),
            span,
        ));
    }
    let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
    let mut stmts = Vec::new();
    for (i, f) in funcs.iter().enumerate() {
        let funcref = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&class)),
            field: f.clone(),
            null_safe: false });
        let call = make_call(
            "table_set",
            vec![
                Expression::int(table_index),
                Expression::int(offset + i as i64),
                funcref,
            ],
            span,
        );
        stmts.push(Statement::new(StmtKind::Expr(call)));
    }
    Ok(Statement::with_span(StmtKind::Block(stmts), span))
}

/// The element storage type of an `array_type` (`(array i8)` → `"i8"`), found
/// as the first `packed_type`/`val_type` under it (descends through `field_def`/
/// `storage_type`/`mut`).
fn array_elem_type(pair: &Pair<Rule>) -> Option<String> {
    for c in pair.clone().into_inner() {
        // Numeric/packed storage (drives sign-extension) OR a ref element type
        // (its text carries `$t`, which drives the typed-null default fill).
        if matches!(
            c.as_rule(),
            Rule::packed_type | Rule::val_type | Rule::ref_val_type
        ) {
            return Some(c.as_str().to_string());
        }
        if let Some(t) = array_elem_type(&c) {
            return Some(t);
        }
    }
    None
}

/// The storage type of one `field_def` (`(field i8)` → `"i8"`, `(field (mut
/// f64))` → `"f64"`, ref fields → the ref type text). Reuses the same
/// packed/val-type search as `array_elem_type`.
fn field_storage_type(field_def: &Pair<Rule>) -> String {
    array_elem_type(field_def).unwrap_or_else(|| {
        // Non-numeric storage (a ref type) — record its text so defaults treat
        // it as a ref (null) rather than a number.
        for c in field_def.clone().into_inner() {
            if matches!(c.as_rule(), Rule::storage_type | Rule::ref_val_type) {
                return c.as_str().to_string();
            }
        }
        String::new()
    })
}

/// The WASM default value for a field storage type: `0` for ints (incl. packed
/// i8/i16), `0.0` for floats, `null` for ref types (`struct.new_default`).
fn default_value_for_storage_type(ty: &str) -> Expression {
    match ty {
        "i8" | "i16" | "i32" | "i64" => Expression::int(0),
        "f32" | "f64" => Expression::float(0.0),
        // A concrete `(ref null $t)` field/element defaults to a WASM GC typed
        // null so an accessor on the defaulted ref traps per spec.
        s if s.contains('$') => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__wast_typed_null")),
            args: vec![],
            optional: false }),
        _ => Expression::null() }
}

/// The ordered field storage types of a `struct_type`/`struct_subtype` body.
fn struct_field_types(composite_inner: &Pair<Rule>) -> Vec<String> {
    composite_inner
        .clone()
        .into_inner()
        .filter(|p| p.as_rule() == Rule::field_def)
        .map(|f| field_storage_type(&f))
        .collect()
}

/// Resolve a table reference (`$t1` name or a numeric index) to its table index.
fn resolve_table_index(name: &str) -> i64 {
    TABLE_NAME_INDEX
        .with(|m| m.borrow().get(name).copied())
        .map(|i| i as i64)
        .or_else(|| name.parse::<i64>().ok())
        .unwrap_or(0)
}

fn walk_data_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut memory_index: u32 = 0;
    let mut offset: Option<Expression> = None;
    let mut bytes: Vec<u8> = Vec::new();
    let mut labels = LabelStack::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            // `(memory idx)(offset …)`, `(offset …)`, or the abbreviated
            // `(i32.const N)` single-instruction offset — all active segments.
            // Absent entirely → passive segment (offset stays None).
            Rule::data_mode => {
                for m in child.into_inner() {
                    match m.as_rule() {
                        Rule::index => {
                            if let Some(i) = m.into_inner().next() {
                                if i.as_rule() == Rule::integer {
                                    memory_index = parse_wat_u64(i.as_str()) as u32;
                                } else {
                                    // `(memory $m2)` — resolve the name to its
                                    // declaration index.
                                    let name = i.as_str().trim_start_matches('$');
                                    memory_index = MEMORY_NAME_INDEX
                                        .with(|f| f.borrow().get(name).copied())
                                        .unwrap_or(0)
                                        as u32;
                                }
                            }
                        }
                        Rule::instr => offset = Some(walk_instr_as_expr(m, &mut labels)?),
                        Rule::folded_instr => {
                            let sp = to_span(&m);
                            offset = Some(walk_folded_instr_as_expr(m, sp, &mut labels)?);
                        }
                        _ => {}
                    }
                }
            }
            Rule::string => bytes.extend(decode_wat_data_string(child.as_str())),
            _ => {}
        }
    }
    Ok(Statement::with_span(
        StmtKind::DataSegment {
            memory_index,
            offset,
            bytes },
        span,
    ))
}

fn parse_wat_u64(s: &str) -> u64 {
    let s = s.trim().replace('_', "");
    if let Some(hex) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
        u64::from_str_radix(hex, 16).unwrap_or(0)
    } else {
        s.parse::<u64>().unwrap_or(0)
    }
}

/// Decode a WAT data-string literal into raw bytes. Data strings differ from
/// text strings: `\HH` (two hex digits) is an arbitrary byte — the dominant
/// form in data segments — alongside `\n \t \r \\ \" \'` and `\u{…}`.
fn decode_wat_data_string(s: &str) -> Vec<u8> {
    let s = s.trim();
    let inner = if s.len() >= 2 && s.starts_with('"') && s.ends_with('"') {
        &s[1..s.len() - 1]
    } else {
        s
    };
    let bytes = inner.as_bytes();
    let hex_val = |c: u8| -> u8 {
        match c {
            b'0'..=b'9' => c - b'0',
            b'a'..=b'f' => c - b'a' + 10,
            b'A'..=b'F' => c - b'A' + 10,
            _ => 0 }
    };
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'\\' || i + 1 >= bytes.len() {
            out.push(bytes[i]);
            i += 1;
            continue;
        }
        match bytes[i + 1] {
            b'n' => {
                out.push(b'\n');
                i += 2;
            }
            b't' => {
                out.push(b'\t');
                i += 2;
            }
            b'r' => {
                out.push(b'\r');
                i += 2;
            }
            b'\\' => {
                out.push(b'\\');
                i += 2;
            }
            b'"' => {
                out.push(b'"');
                i += 2;
            }
            b'\'' => {
                out.push(b'\'');
                i += 2;
            }
            b'u' if i + 2 < bytes.len() && bytes[i + 2] == b'{' => {
                if let Some(close_rel) = inner[i + 3..].find('}') {
                    let hex = &inner[i + 3..i + 3 + close_rel];
                    if let Ok(cp) = u32::from_str_radix(hex, 16) {
                        if let Some(ch) = char::from_u32(cp) {
                            let mut buf = [0u8; 4];
                            out.extend_from_slice(ch.encode_utf8(&mut buf).as_bytes());
                        }
                    }
                    i += 3 + close_rel + 1;
                } else {
                    i += 2;
                }
            }
            c if c.is_ascii_hexdigit()
                && i + 2 < bytes.len()
                && bytes[i + 2].is_ascii_hexdigit() =>
            {
                out.push((hex_val(c) << 4) | hex_val(bytes[i + 2]));
                i += 3;
            }
            c => {
                out.push(c);
                i += 2;
            }
        }
    }
    out
}

// ── WAST script commands ──────────────────────────────────────────────────────

fn walk_invoke_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut args: Vec<Expression> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {}
            Rule::string => {
                if func_name.is_empty() {
                    func_name = unquote(child.as_str());
                }
            }
            Rule::expr => args.push(walk_const_expr(child)?),
            _ => {}
        }
    }
    // Resolve the exported name to the module class's static method so the call
    // actually reaches the function (exports are `Class.method`).
    let callee = match EXPORT_FUNC_MAP.with(|m| m.borrow().get(&func_name).cloned()) {
        Some(method) => {
            let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
            Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::ident(&class)),
                    field: method,
                    null_safe: false },
                span,
            )
        }
        None => Expression::ident(&func_name) };
    Ok(Statement::with_span(
        StmtKind::Expr(Expression::with_span(
            ExprKind::Call {
                callee: Box::new(callee),
                args: args.into_iter().map(Argument::positional).collect(),
                optional: false },
            span,
        )),
        span,
    ))
}

/// `(assert_return (invoke "f" …) (i32.const 42))`.
///
/// A script directive is part of the LANGUAGE, not a call into a test harness.
/// It lowers to ordinary code — run the action, compare, throw on mismatch — so
/// `vybex file.wast` works on its own. Routing it to `vybe:wast:assert_return`
/// meant the only implementation lived in `languages/wast/tests/wast/helpers.rs`
/// and every script died with `Unresolved import` outside `cargo test`.
///
/// The file itself is untouched: `wasmtime wast` runs the same source through
/// its own native directive support.
fn walk_assert_return(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    let mut expected: Vec<Expression> = Vec::new();
    let mut expects_nan = Vec::new();
    let mut expects_v128 = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action => action_expr = Some(walk_action(child)?),
            Rule::result_val => {
                // `nan:canonical` / `nan:arithmetic` pin no payload, so they
                // cannot be compared with `==`.
                expects_nan.push(child.as_str().contains("nan"));
                // A v128 is a VECTOR: scalar `!=` does not compare its lanes,
                // so it reports "different" for two identical vectors and every
                // v128-returning assert fails regardless of its lanes.
                expects_v128.push(child.as_str().contains("v128"));
                expected.push(walk_const_expr(child)?);
            }
            _ => {}
        }
    }
    let Some(action) = action_expr else {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    };

    if expected.len() == 1 {
        let want = expected.pop().expect("checked len");
        let cond = if expects_nan[0] {
            // NaN is the only value that differs from itself, so "equals
            // itself" is precisely the failure case.
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(action.clone()),
                    right: Box::new(action) },
                span,
            )
        } else if expects_v128[0] {
            // Lane-wise: `i8x16.eq` yields all-ones in each byte lane that
            // matches, and `i8x16.all_true` is 1 only when every lane did. So
            // the FAILURE condition is that result being 0. Comparing at i8x16
            // is shape-independent — it checks all 16 bytes, so it is correct
            // for i32x4 / f64x2 / any other interpretation of the same bits.
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(make_call(
                        "i8x16.all_true",
                        vec![make_call("i8x16.eq", vec![action, want], span)],
                        span,
                    )),
                    right: Box::new(Expression::int(0)) },
                span,
            )
        } else {
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(action),
                    right: Box::new(want) },
                span,
            )
        };
        let throw = Statement::with_span(
            StmtKind::Throw {
                expr: Some(Expression::string("assert_return failed")),
                cause: None },
            span,
        );
        return Ok(Statement::with_span(
            StmtKind::If {
                cond,
                then_body: vec![throw],
                elifs: Vec::new(),
                else_body: None },
            span,
        ));
    }

    // Multi-value results arrive packed as one array and zero-value actions
    // have nothing to compare; both still go through the builtin.
    let mut args = vec![action];
    args.extend(expected);
    Ok(Statement::with_span(
        StmtKind::Expr(make_call("__wast_assert_return", args, span)),
        span,
    ))
}

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
    if let Some(a) = action_expr {
        args.push(a);
    }
    args.push(Expression::string(&message));
    Ok(Statement::with_span(
        StmtKind::Expr(make_call("__wast_assert_trap", args, span)),
        span,
    ))
}

fn walk_assert_generic(pair: Pair<Rule>, fn_name: &str) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut message = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            message = unquote(child.as_str());
        }
    }
    Ok(Statement::with_span(
        StmtKind::Expr(make_call(fn_name, vec![Expression::string(&message)], span)),
        span,
    ))
}

fn walk_register_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            name = unquote(child.as_str());
            break;
        }
    }
    Ok(Statement::with_span(
        StmtKind::Expr(make_call(
            "__wasm_register",
            vec![Expression::string(&name)],
            span,
        )),
        span,
    ))
}

fn walk_get_cmd(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            name = unquote(child.as_str());
            break;
        }
    }
    Ok(Statement::with_span(
        StmtKind::Expr(make_call(
            "__wasm_get",
            vec![Expression::string(&name)],
            span,
        )),
        span,
    ))
}

fn walk_action(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty action")?;
    match inner.as_rule() {
        Rule::invoke_cmd => {
            let stmt = walk_invoke_cmd(inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()) }
        }
        Rule::get_cmd => {
            let stmt = walk_get_cmd(inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()) }
        }
        _ => Ok(Expression::null()) }
}

fn walk_const_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let raw_text = pair.as_str().to_string();
    // A `(v128.const <lane> l0 l1 …)` expected result carries a `val_lane_type`
    // plus its lane integers. Reconstruct the SAME lowering the actual side uses
    // (`v128.const` → fallback `make_call("v128_const", [lane, l0, l1, …])` in
    // `map_instr_to_ast`) so the expected v128 compares byte-identically to the
    // computed one — otherwise the scalar path below would collapse the whole
    // vector to its first lane.
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if let Some(lane) = children.iter().find(|c| c.as_rule() == Rule::val_lane_type) {
        let mut args = vec![Expression::string(lane.as_str())];
        for c in &children {
            if c.as_rule() == Rule::integer {
                args.push(parse_integer(c.as_str()));
            }
        }
        return Ok(make_call("v128_const", args, span));
    }
    // `ref.func`/`ref.extern` keywords are literal tokens (not captured rules),
    // so inspect the raw text. Bare `(ref.func)` is the spec's abstract pattern
    // "any non-null funcref" → a sentinel the assert harness matches against any
    // funcref. `(ref.extern N)` carries its integer payload (caught below).
    if raw_text.contains("ref.func") {
        return Ok(Expression::string("__wast_any_funcref"));
    }
    for child in children {
        match child.as_rule() {
            Rule::integer => return Ok(parse_integer(child.as_str())),
            Rule::float => return Ok(parse_float(child.as_str())),
            Rule::string => return Ok(Expression::string(&unquote(child.as_str()))),
            _ => {}
        }
    }
    Ok(Expression::null())
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn make_call(name: &str, args: Vec<Expression>, span: Span) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false },
        span,
    )
}

fn bin_op(mut args: Vec<Expression>, op: BinOp, span: Span) -> Result<Expression, String> {
    let right = if args.len() >= 2 {
        args.remove(1)
    } else {
        Expression::int(0)
    };
    let left = args.into_iter().next().unwrap_or(Expression::int(0));
    Ok(Expression::with_span(
        ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right) },
        span,
    ))
}

/// WAT functions implicitly return the last value left on the stack.
fn apply_implicit_return(body: &mut Vec<Statement>) {
    if body.is_empty() {
        return;
    }
    if let Some(last) = body.last_mut() {
        if let StmtKind::Expr(ref e) = last.kind.clone() {
            if let ExprKind::Call { ref callee, .. } = e.kind {
                if let ExprKind::Ident(ref n) = callee.kind {
                    if n == "__wasm_return" {
                        return;
                    }
                }
            }
            last.kind = StmtKind::Return(Some(e.clone()));
        }
    }
}

/// A `(result t1 t2 …)` (N ≥ 2) function implicitly returns the top N stack
/// values. Gather the trailing N flushed value-statements into one uniform
/// tuple `return`, which the shared compiler's multi-value ABI
/// (`uniform_tuple_return_arity`) recognises → `result_arity = N`, pushing the
/// elements unpacked for the caller to destructure. If the body doesn't end in
/// N contiguous value-statements (e.g. it always branched out via an explicit
/// `return`, already tuple-shaped), we leave it untouched.
fn apply_multi_value_return(body: &mut Vec<Statement>, n: usize) {
    let mut idxs: Vec<usize> = Vec::with_capacity(n);
    for (i, s) in body.iter().enumerate().rev() {
        if matches!(s.kind, StmtKind::Expr(_)) {
            idxs.push(i);
            if idxs.len() == n {
                break;
            }
        } else {
            break;
        }
    }
    if idxs.len() != n {
        return;
    }
    idxs.reverse(); // ascending = stack bottom-to-top → tuple element order
    let elems: Vec<Expression> = idxs
        .iter()
        .map(|&i| match &body[i].kind {
            StmtKind::Expr(e) => e.clone(),
            _ => unreachable!() })
        .collect();
    // The N statements are contiguous at the tail; drop them and append the
    // single tuple return.
    body.truncate(idxs[0]);
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Tuple(elems),
    )))));
}

/// Build a multi-value `return` of the top `n` stack values as a tuple (used by
/// an explicit `return` inside a multi-value function). `temps[0]` ← deepest.
fn multi_value_return_stmt(stack: &mut Vec<Expression>, n: usize, span: Span) -> Statement {
    let avail = n.min(stack.len());
    let elems: Vec<Expression> = stack.split_off(stack.len() - avail);
    Statement::with_span(
        StmtKind::Return(Some(Expression::new(ExprKind::Tuple(elems)))),
        span,
    )
}

fn parse_integer(s: &str) -> Expression {
    let s = s.trim().replace('_', "");
    let (neg, digits) = if s.starts_with("-0x") || s.starts_with("-0X") {
        (true, &s[3..])
    } else if s.starts_with("0x")
        || s.starts_with("0X")
        || s.starts_with("+0x")
        || s.starts_with("+0X")
    {
        (false, if s.starts_with('+') { &s[3..] } else { &s[2..] })
    } else {
        // Decimal: fall back to u64 (then reinterpret) for values above i64::MAX
        // such as an unsigned 64-bit literal written in decimal.
        let v = s
            .parse::<i64>()
            .or_else(|_| s.parse::<u64>().map(|u| u as i64))
            .unwrap_or(0);
        return Expression::int(v);
    };
    // Hex: a 64-bit pattern like 0x8000000000000000 exceeds i64::MAX, so parse as
    // u64 and reinterpret to the signed value it denotes.
    let v = i64::from_str_radix(digits, 16)
        .or_else(|_| u64::from_str_radix(digits, 16).map(|u| u as i64))
        .unwrap_or(0);
    Expression::int(if neg { v.wrapping_neg() } else { v })
}

fn parse_float(s: &str) -> Expression {
    let s = s.trim();
    match s {
        "inf" | "+inf" => Expression::float(f64::INFINITY),
        "-inf" => Expression::float(f64::NEG_INFINITY),
        // All NaN forms — plain, `nan:0x…`, `nan:canonical`, `nan:arithmetic`
        // (payload/kind is not observable through an f64 in this VM).
        _ if s.contains("nan") => Expression::float(f64::NAN),
        _ => {
            let cleaned = s.replace('_', "");
            if cleaned.contains("0x") || cleaned.contains("0X") {
                if let Some(v) = parse_hex_float(&cleaned) {
                    return Expression::float(v);
                }
            }
            Expression::float(cleaned.parse::<f64>().unwrap_or(0.0))
        }
    }
}

/// Parse a WAT hex float like `0x1.8p1` (= 1.5 × 2¹ = 3.0). Rust's `f64::parse`
/// rejects hex floats, so evaluate the mantissa (hex int + hex fraction) and
/// scale by the binary `p` exponent.
fn parse_hex_float(s: &str) -> Option<f64> {
    let (neg, rest) = match s.strip_prefix('-') {
        Some(r) => (true, r),
        None => (false, s.strip_prefix('+').unwrap_or(s)) };
    let rest = rest
        .strip_prefix("0x")
        .or_else(|| rest.strip_prefix("0X"))?;
    let (mantissa, exp) = match rest.find(['p', 'P']) {
        Some(i) => (&rest[..i], rest[i + 1..].parse::<i32>().ok()?),
        None => (rest, 0) };
    let (int_part, frac_part) = match mantissa.find('.') {
        Some(i) => (&mantissa[..i], &mantissa[i + 1..]),
        None => (mantissa, "") };
    let mut value = 0.0f64;
    for c in int_part.chars() {
        value = value * 16.0 + c.to_digit(16)? as f64;
    }
    let mut scale = 1.0 / 16.0;
    for c in frac_part.chars() {
        value += c.to_digit(16)? as f64 * scale;
        scale /= 16.0;
    }
    value *= 2f64.powi(exp);
    Some(if neg { -value } else { value })
}

fn unquote(s: &str) -> String {
    let s = s.trim();
    if s.len() >= 2 && s.starts_with('"') && s.ends_with('"') {
        s[1..s.len() - 1]
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
    let end = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32,
        start_col: start.1 as u32,
        end_line: end.0 as u32,
        end_col: end.1 as u32 }
}

/// Peek the plain-instruction keyword of an `instr`/`plain_instr` pair without
/// consuming it. Returns None for folded instructions (which carry no linear
/// `block`/`loop`/`if`/`else`/`end` tokens).
/// The folded instructions a `block`/`loop`/`if` OPENER swallowed.
///
/// `plain_instr = instr_name ~ instr_arg*`, and `instr_arg` accepts a
/// `folded_instr`. So in the UNFOLDED form
///
/// ```wat
/// if (result i32)
///   (i32.const 5)
/// else
/// ```
///
/// the `(i32.const 5)` parses as an ARGUMENT of the `if`, not as the first
/// instruction of its branch. `find_matching_end` then slices an EMPTY branch
/// body and the branch's result temp keeps its `null` initialiser.
///
/// Measured 2026-08-06: the folded form yields `null` where the plain form
/// yields `5` — for EVERY instruction, silently, exit 0. It looked like an
/// `unreachable` bug because a folded `(unreachable)` stopped trapping, but the
/// instruction never reached any lowering at all.
///
/// Give them back to the branch. `block_type` (`(result i32)`) and `id`
/// (`$label`) are genuinely the opener's own arguments and stay put; only
/// folded instructions move.
fn opener_folded_instrs<'a>(pair: &Pair<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    let inner = if pair.as_rule() == Rule::instr {
        match pair.clone().into_inner().next() {
            Some(p) => p,
            None => return Vec::new() }
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return Vec::new();
    }
    inner
        .into_inner()
        .filter(|c| c.as_rule() == Rule::instr_arg)
        .filter_map(|arg| arg.into_inner().find(|x| x.as_rule() == Rule::folded_instr))
        .collect()
}

fn peek_plain_name(pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    if inner.as_rule() == Rule::plain_instr {
        inner
            .into_inner()
            .find(|c| c.as_rule() == Rule::instr_name)
            .map(|c| c.as_str().to_string())
    } else {
        None
    }
}

/// The `$id` label immediately following a `block`/`loop`/`if` keyword, if any.
fn peek_plain_label(pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return None;
    }
    inner
        .into_inner()
        .filter_map(|c| {
            if c.as_rule() == Rule::instr_arg {
                c.into_inner().next()
            } else {
                None
            }
        })
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
}

/// The `(type $sig)` type name from a `call_indirect`/`return_call_indirect`
/// opener — its `block_type` immediate wraps a `(type index)` typeuse. Returns
/// the type index text (`$`-stripped) so its param count gives the call arity.
fn peek_typeuse_index(pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return None;
    }
    for c in inner.into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        if let Some(bt) = c.into_inner().next() {
            if bt.as_rule() == Rule::block_type {
                for x in bt.into_inner() {
                    if x.as_rule() == Rule::index {
                        return Some(x.as_str().trim_start_matches('$').to_string());
                    }
                }
            }
        }
    }
    None
}

/// The optional table reference of `call_indirect $t (type $sig)` — the first
/// bare `id`/`index` arg (the `(type …)` sig is a `block_type`, so it's skipped).
/// Returns None for the default-table form `call_indirect (type $sig)`.
fn peek_call_indirect_table(pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return None;
    }
    for c in inner.into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        if let Some(a) = c.into_inner().next() {
            if matches!(a.as_rule(), Rule::index | Rule::id) {
                return Some(a.as_str().trim_start_matches('$').to_string());
            }
        }
    }
    None
}

/// Given the index of an unfolded `block`/`loop`/`if` opener, find the matching
/// `end` (respecting nesting) and, for `if`, the `else` at the same depth.
fn find_matching_end(
    pairs: &[Pair<Rule>],
    opener: usize,
) -> Result<(Option<usize>, usize), String> {
    let mut depth = 1usize;
    let mut else_idx: Option<usize> = None;
    let mut j = opener + 1;
    while j < pairs.len() {
        if let Some(kw) = peek_plain_name(&pairs[j]) {
            match kw.as_str() {
                "block" | "loop" | "if" | "try" => depth += 1,
                "else" if depth == 1 => else_idx = Some(j),
                // A legacy `delegate N` closes its `try` with no `end`; count it
                // as a closer so a delegate-try nested here doesn't unbalance us.
                "delegate" => depth -= 1,
                "end" => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok((else_idx, j));
                    }
                }
                _ => {}
            }
        }
        j += 1;
    }
    Err("unterminated block/loop/if (missing end)".to_string())
}

fn fold_instructions(
    pairs: Vec<Pair<Rule>>,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    fold_instructions_seeded(pairs, labels, Vec::new())
}

/// Like `fold_instructions`, but the value stack starts pre-loaded with `seed`
/// (bottom-to-top). Used to thread WASM `block (param …)` inputs into the body.
fn fold_instructions_seeded(
    pairs: Vec<Pair<Rule>>,
    labels: &mut LabelStack,
    seed: Vec<Expression>,
) -> Result<Vec<Statement>, String> {
    let mut stack: Vec<Expression> = seed;
    let mut statements: Vec<Statement> = Vec::new();

    let mut i = 0;
    while i < pairs.len() {
        // ── Unfolded structured control: block/loop/if … else … end ──────────
        // These arrive as flat `plain_instr` tokens; group them into the same
        // Labeled/Block/While/If statements the folded S-expr forms produce.
        if let Some(kw) = peek_plain_name(&pairs[i]) {
            match kw.as_str() {
                "block" | "loop" | "if" => {
                    let span = to_span(&pairs[i]);
                    let label = peek_plain_label(&pairs[i]);
                    let (else_idx, end_idx) = find_matching_end(&pairs, i)?;

                    if kw == "if" {
                        let result_temps: Vec<String> = (0..peek_block_result_count(&pairs[i]))
                            .map(|_| fresh_result_temp())
                            .collect();
                        // Condition is the value on top of the stack; below it sit
                        // the `(param …)` block-type inputs, which WASM threads into
                        // BOTH branch bodies — split them off to seed each fold
                        // (any values below the params are pending side effects).
                        let cond = stack.pop().unwrap_or(Expression::bool(false));
                        let param_count = peek_block_param_count(&pairs[i]);
                        let seed = if param_count > 0 && stack.len() >= param_count {
                            stack.split_off(stack.len() - param_count)
                        } else {
                            Vec::new()
                        };
                        preserve_stack_across_block(&mut stack, &mut statements);
                        let then_end = else_idx.unwrap_or(end_idx);
                        // Anything the opener swallowed belongs to the THEN
                        // branch, ahead of what follows the opener token.
                        let mut then_pairs: Vec<Pair<Rule>> = opener_folded_instrs(&pairs[i]);
                        then_pairs.extend(pairs[i + 1..then_end].iter().cloned());
                        labels.push(label.clone(), LabelKind::Block, Vec::new());
                        let mut then_body =
                            fold_instructions_seeded(then_pairs, labels, seed.clone())?;
                        let mut else_body = if let Some(ei) = else_idx {
                            // `else` is a `plain_instr` too, so it swallows the
                            // first folded instruction of ITS branch the same
                            // way the opener does.
                            let mut else_pairs: Vec<Pair<Rule>> =
                                opener_folded_instrs(&pairs[ei]);
                            else_pairs.extend(pairs[ei + 1..end_idx].iter().cloned());
                            Some(fold_instructions_seeded(else_pairs, labels, seed)?)
                        } else {
                            None
                        };
                        labels.pop();
                        // A `(result …)` if yields N values: capture each branch's
                        // trailing N values in N temps and leave them on the stack.
                        if !result_temps.is_empty() {
                            for tmp in &result_temps {
                                statements.push(Statement::new(StmtKind::VarDecl {
                                    declarations: vec![VarDeclarator {
                                        pattern: BindingPattern::Ident(tmp.clone()),
                                        type_hint: None,
                                        init: Some(Expression::null()),
                                        array_bounds: None,
                                        with_events: false }],
                                    kind: VarDeclKind::Let }));
                            }
                            assign_last_n_exprs_to(&mut then_body, &result_temps);
                            if let Some(eb) = else_body.as_mut() {
                                assign_last_n_exprs_to(eb, &result_temps);
                            }
                            statements.push(Statement::with_span(
                                StmtKind::If {
                                    cond,
                                    then_body,
                                    else_body,
                                    elifs: Vec::new() },
                                span,
                            ));
                            for tmp in &result_temps {
                                stack.push(Expression::ident(tmp));
                            }
                        } else {
                            statements.push(Statement::with_span(
                                StmtKind::If {
                                    cond,
                                    then_body,
                                    else_body,
                                    elifs: Vec::new() },
                                span,
                            ));
                        }
                    } else {
                        // block / loop take no condition. Pop the block's
                        // param values off the top to seed the body, then
                        // sequence any remaining pending side effects.
                        let param_count = peek_block_param_count(&pairs[i]);
                        let seed_vals = if param_count > 0 && stack.len() >= param_count {
                            stack.split_off(stack.len() - param_count)
                        } else {
                            Vec::new()
                        };
                        preserve_stack_across_block(&mut stack, &mut statements);
                        // Same for `block`/`loop`: a folded instruction written
                        // first in the body parses as an opener argument.
                        let mut body_pairs: Vec<Pair<Rule>> = opener_folded_instrs(&pairs[i]);
                        body_pairs.extend(pairs[i + 1..end_idx].iter().cloned());
                        // A `loop (param …)` threads its operand-stack params
                        // across iterations. Model each with a synthetic local:
                        // initialise it from the entry value, let the body read it
                        // (its seed), and have every `br` back to the loop assign
                        // the next iteration's value into it (see the `br` arm).
                        // This makes the loop a real `while(true)` rather than the
                        // one-shot block a param-less lowering would force.
                        let loop_has_param = kw == "loop" && peek_opener_has_param(&pairs[i]);
                        let param_temps: Vec<String> = if loop_has_param {
                            (0..param_count).map(|_| fresh_result_temp()).collect()
                        } else {
                            Vec::new()
                        };
                        let seed: Vec<Expression> = if loop_has_param {
                            for (k, tmp) in param_temps.iter().enumerate() {
                                let init = seed_vals
                                    .get(k)
                                    .cloned()
                                    .unwrap_or_else(|| Expression::int(0));
                                statements.push(Statement::new(StmtKind::VarDecl {
                                    declarations: vec![VarDeclarator {
                                        pattern: BindingPattern::Ident(tmp.clone()),
                                        type_hint: None,
                                        init: Some(init),
                                        array_bounds: None,
                                        with_events: false }],
                                    kind: VarDeclKind::Let }));
                            }
                            param_temps.iter().map(|t| Expression::ident(t)).collect()
                        } else {
                            seed_vals
                        };
                        let kind = if kw == "block" {
                            LabelKind::Block
                        } else {
                            LabelKind::Loop
                        };
                        // A `(result …)` block/loop yields N values: `br` to it
                        // carries the top N stack values into N temps, and the
                        // fall-through assigns the same temps; the temps are left
                        // on the stack. N == 1 is the single-value baseline.
                        let result_temps: Vec<String> = (0..peek_block_result_count(&pairs[i]))
                            .map(|_| fresh_result_temp())
                            .collect();
                        for tmp in &result_temps {
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(tmp.clone()),
                                    type_hint: None,
                                    init: Some(Expression::null()),
                                    array_bounds: None,
                                    with_events: false }],
                                kind: VarDeclKind::Let }));
                        }
                        let effective = labels.push(label.clone(), kind, result_temps.clone());
                        labels.set_last_param_temps(param_temps.clone());
                        let mut body = fold_instructions_seeded(body_pairs, labels, seed)?;
                        labels.pop();
                        // Capture the fall-through values (unreachable if the body
                        // always branches out, which is why it's safe to append).
                        assign_last_n_exprs_to(&mut body, &result_temps);
                        let inner_stmt = if kw == "block" {
                            Statement::with_span(StmtKind::Block(body), span)
                        } else {
                            // A WASM loop exits when control falls off its end;
                            // `while (true)` needs an explicit break to match.
                            body.push(Statement::with_span(
                                StmtKind::Break(BreakTarget::Implicit),
                                span,
                            ));
                            Statement::with_span(
                                StmtKind::While {
                                    cond: Expression::bool(true),
                                    body,
                                    else_body: None },
                                span,
                            )
                        };
                        statements.push(Statement::with_span(
                            StmtKind::Labeled {
                                label: effective,
                                body: Box::new(inner_stmt) },
                            span,
                        ));
                        for tmp in &result_temps {
                            stack.push(Expression::ident(tmp));
                        }
                    }
                    i = end_idx + 1;
                    continue;
                }
                // ── throw $tag: raise with the top `arity` stack values ─────
                "throw" => {
                    let span = to_span(&pairs[i]);
                    let tag = peek_plain_label(&pairs[i]).unwrap_or_default();
                    let arity = tag_arity(&tag) as usize;
                    let n = arity.min(stack.len());
                    let args: Vec<Expression> = stack.split_off(stack.len() - n);
                    statements.push(Statement::with_span(
                        StmtKind::WasmThrow { tag, args },
                        span,
                    ));
                    i += 1;
                    continue;
                }
                // ── throw_ref: re-raise an `exnref` taken from the stack ────────
                // (canonical WASM 3.0; supersedes legacy `rethrow N`). Bind the
                // exnref operand to a local and reuse the `WasmRethrow` lowering,
                // which reads that local and emits `Op::THROW_REF`.
                "throw_ref" => {
                    let span = to_span(&pairs[i]);
                    let exnref_expr = stack.pop().unwrap_or_else(Expression::null);
                    let exnref_local = match &exnref_expr.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => {
                            let tmp = fresh_result_temp();
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(tmp.clone()),
                                    type_hint: None,
                                    init: Some(exnref_expr),
                                    array_bounds: None,
                                    with_events: false }],
                                kind: VarDeclKind::Let }));
                            tmp
                        }
                    };
                    statements.push(Statement::with_span(
                        StmtKind::WasmRethrow { exnref_local },
                        span,
                    ));
                    i += 1;
                    continue;
                }
                // Unfolded `br_on_null $L` / `br_on_non_null $L`: the ref is
                // already on the value stack (from a prior instruction) and the
                // label is the instr_arg. Reuse the folded lowering (which pops
                // the ref and reads the label) so a proper structured branch is
                // emitted — the generic path would emit the label where the VM
                // expects a relative offset, misaligning the stream.
                "br_on_null" | "br_on_non_null" => {
                    let span = to_span(&pairs[i]);
                    let is_non_null = kw == "br_on_non_null";
                    emit_folded_br_on_null(
                        pairs[i].clone(),
                        is_non_null,
                        span,
                        labels,
                        &mut statements,
                        &mut stack,
                    )?;
                    i += 1;
                    continue;
                }
                // call_indirect (type $sig): call a funcref via a table. Supply
                // the argc (from the sig's params) + tableidx immediates and
                // the stack operands (spec order: call args then the table
                // index on top). Handled here (not the generic path) so the
                // `(type $sig)` — dropped by the generic arg walk — is read.
                "call_indirect" | "return_call_indirect" => {
                    let span = to_span(&pairs[i]);
                    let sig = peek_typeuse_index(&pairs[i]);
                    let argc = sig
                        .as_ref()
                        .and_then(|n| TYPE_FUNC_PARAMS.with(|m| m.borrow().get(n).copied()))
                        .unwrap_or(0) as usize;
                    // Expected result count — the other half of the type shape
                    // the VM checks the funcref against (traps on mismatch).
                    let expected_results = sig
                        .as_ref()
                        .and_then(|n| TYPE_FUNC_RESULTS.with(|m| m.borrow().get(n).copied()))
                        .unwrap_or(0) as usize;
                    // Optional table reference `call_indirect $t (type $sig)`
                    // dispatches through a NAMED table (default table 0).
                    let tableidx = peek_call_indirect_table(&pairs[i])
                        .map(|t| resolve_table_index(&t) as usize)
                        .unwrap_or(0);
                    let n = (argc + 1).min(stack.len());
                    let operands: Vec<Expression> = stack.split_off(stack.len() - n);
                    let mut call_args = vec![
                        Expression::int(argc as i64),
                        Expression::int(tableidx as i64),
                        Expression::int(expected_results as i64),
                    ];
                    call_args.extend(operands);
                    // `return_call_indirect` is the tail-call form: it emits the
                    // frame-reusing `RETURN_CALL_INDIRECT` opcode (same immediate
                    // + type check as `call_indirect`) and diverges, rather than
                    // a `call` + `return` (which would grow the stack).
                    if kw == "return_call_indirect" {
                        let call = make_call("return_call_indirect", call_args, span);
                        statements.push(Statement::with_span(StmtKind::Expr(call), span));
                    } else {
                        stack.push(make_call("call_indirect", call_args, span));
                    }
                    i += 1;
                    continue;
                }
                "end" | "else" => {
                    // Stray delimiter (already consumed by find_matching_end for
                    // real openers) — skip defensively.
                    i += 1;
                    continue;
                }
                _ => {}
            }
        }

        let pair = pairs[i].clone();
        i += 1;
        let span = to_span(&pair);
        let inner = if pair.as_rule() == Rule::instr {
            pair.into_inner().next().ok_or("Empty instr")?
        } else {
            pair
        };

        match inner.as_rule() {
            Rule::folded_instr => {
                // Structured control (`block`/`loop`) and value-carrying branches
                // (`br_on_null`/`br_on_non_null`/`return`) need STATEMENT lowering
                // — `walk_folded_core` returns only an expression and discards a
                // folded block's body. Route those to dedicated handlers; the rest
                // stay on the expression path.
                let head = folded_instr_head(&inner);
                match head.as_str() {
                    "block" => {
                        emit_folded_block(inner, false, span, labels, &mut statements, &mut stack)?;
                    }
                    "loop" => {
                        emit_folded_block(inner, true, span, labels, &mut statements, &mut stack)?;
                    }
                    "try_table" => {
                        emit_folded_try_table(inner, span, labels, &mut statements, &mut stack)?;
                    }
                    "br_on_null" => {
                        emit_folded_br_on_null(
                            inner,
                            false,
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        )?;
                    }
                    "br_on_non_null" => {
                        emit_folded_br_on_null(
                            inner,
                            true,
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        )?;
                    }
                    "return" => {
                        emit_folded_return(inner, span, labels, &mut statements, &mut stack)?;
                    }
                    // `(drop (br_on_null $L …))` — the branch's fall-through ref is
                    // discarded. Handle the branch (which leaves the non-null ref
                    // on the stack), then pop it for the `drop`.
                    "drop" => {
                        // A folded operand nests as `instr_arg → folded_instr`
                        // (the grammar's `instr_arg*` greedily matches nested
                        // folded instrs before `instr*`).
                        let nested = inner
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::instr_arg)
                            .find_map(|arg| {
                                arg.into_inner().find(|x| x.as_rule() == Rule::folded_instr)
                            })
                            .or_else(|| {
                                inner
                                    .clone()
                                    .into_inner()
                                    .find(|c| c.as_rule() == Rule::instr)
                                    .and_then(|i| i.into_inner().next())
                            });
                        let nested_head =
                            nested.as_ref().map(folded_instr_head).unwrap_or_default();
                        match (nested, nested_head.as_str()) {
                            (Some(op), "br_on_null") => {
                                emit_folded_br_on_null(
                                    op,
                                    false,
                                    span,
                                    labels,
                                    &mut statements,
                                    &mut stack,
                                )?;
                                stack.pop();
                            }
                            (Some(op), "br_on_non_null") => {
                                emit_folded_br_on_null(
                                    op,
                                    true,
                                    span,
                                    labels,
                                    &mut statements,
                                    &mut stack,
                                )?;
                                stack.pop();
                            }
                            _ => {
                                // Ordinary drop: a void instr, emit in program order.
                                let expr = walk_folded_instr_as_expr(inner, span, labels)?;
                                statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                            }
                        }
                    }
                    _ => {
                        // A folded operand nests as `instr` or `instr_arg →
                        // folded/plain_instr`. When a stack-consuming instr has
                        // NONE of its value operands nested (the abbreviated
                        // "flat" style, e.g. `(struct.get $t 0)` after a block
                        // that left the ref on the stack), it takes its operands
                        // from the enclosing stack exactly like the plain form.
                        // Without this the missing operand became null.
                        let has_nested_operand = inner.clone().into_inner().any(|c| {
                            c.as_rule() == Rule::instr
                                || (c.as_rule() == Rule::instr_arg
                                    && matches!(
                                        c.into_inner().next().map(|x| x.as_rule()),
                                        Some(Rule::folded_instr) | Some(Rule::plain_instr)
                                    ))
                        });
                        let immediate_args: Vec<Expression> = inner
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::instr_arg)
                            .map(|c| walk_instr_arg_pair(c, labels))
                            .collect::<Result<_, _>>()?;
                        let arity = get_instruction_arity(&head, &immediate_args);
                        if !has_nested_operand && head != "call" && arity > 0 {
                            let mut args = immediate_args;
                            let pop_count = arity.min(stack.len());
                            let drain_start = stack.len() - pop_count;
                            let popped: Vec<Expression> = stack.drain(drain_start..).collect();
                            args.extend(popped);
                            let expr = map_instr_to_ast(head.clone(), args, span)?;
                            if get_instruction_push_count(&head) > 0 {
                                stack.push(expr);
                            } else {
                                statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                            }
                            continue;
                        }
                        // A fully-folded instr is self-contained (all operands
                        // nested), so statement-vs-stack is purely an ordering
                        // question. Void instructions (`local.set`, stores,
                        // `struct.set`/`array.set`, bulk ops) must run in program
                        // order — deferring them on the value stack lets a later
                        // reader observe the pre-write state. Value producers — and
                        // `call`, whose result count is context-dependent — stay on
                        // the stack for their consumer.
                        let expr = walk_folded_instr_as_expr(inner, span, labels)?;
                        if !head.is_empty()
                            && head != "call"
                            && get_instruction_push_count(&head) == 0
                        {
                            statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                        } else {
                            stack.push(expr);
                        }
                    }
                }
            }
            Rule::plain_instr => {
                let mut name = String::new();
                let mut raw_args = Vec::new();
                for child in inner.clone().into_inner() {
                    match child.as_rule() {
                        Rule::instr_name => name = child.as_str().to_string(),
                        Rule::instr_arg => raw_args.push(child),
                        _ => {}
                    }
                }

                // Multi-memory: peel any leading bare memidx immediate(s) into a
                // `@@mem<N>` name suffix BEFORE parsing the remaining operands, so
                // a real selector is never confused with a greedily-attached
                // folded operand.
                let name = peel_mem_selector(&name, &mut raw_args, labels)?;

                // Parse inline arguments
                let mut args = Vec::new();
                for raw in raw_args {
                    args.push(walk_instr_arg_pair(raw, labels)?);
                }

                // Determine stack arity
                let arity = get_instruction_arity(&name, &args);
                let pop_count = usize::min(arity, stack.len());
                let drain_start = stack.len() - pop_count;
                let popped: Vec<Expression> = stack.drain(drain_start..).collect();

                // Append popped operands to args
                args.extend(popped);

                // Handle control flow instructions that are statements
                match name.as_str() {
                    "nop" => {
                        // Nop does nothing, ignore
                    }
                    // A trap, NOT a throw — see `trap_expr`. This is the LIVE
                    // lowering: the `walk_*_instr_as_stmts` family that also
                    // spelled `unreachable` is dead code.
                    //
                    // Pushed onto the operand STACK as well as emitted, so a
                    // folded `(unreachable)` in value position — `(func $f
                    // (result i32) (unreachable))` — is the trap rather than the
                    // `Expression::null()` it used to be. That null made the
                    // function return normally and the caller exit 0, so any
                    // wast test whose failure path was written that way passed
                    // unconditionally.
                    "unreachable" => {
                        statements.push(Statement::with_span(
                            StmtKind::Expr(trap_expr()),
                            span,
                        ));
                        stack.push(trap_expr());
                    }
                    "return" => {
                        let n = CURRENT_FN_RESULTS.with(|c| *c.borrow());
                        if n >= 2 {
                            // Multi-value function: reraise the top N values as a
                            // uniform tuple (multi-value ABI).
                            statements.push(multi_value_return_stmt(&mut stack, n, span));
                        } else {
                            let val = stack.pop();
                            statements.push(Statement::with_span(StmtKind::Return(val), span));
                        }
                    }
                    // Tail calls: `return_call $f` / `return_call_ref` must
                    // REUSE the frame (WASM tail-call proposal) so unbounded
                    // tail recursion runs in O(1) stack. Reuse the `call`/
                    // `call_ref` lowering to qualify the callee, then emit a
                    // `__wasm_return_call(callee, args…)` which the compiler
                    // lowers to the frame-reusing `Op::RETURN_CALL` (a plain
                    // `return f(args)` would grow the stack and overflow).
                    "return_call" | "return_call_ref" => {
                        let inner = if name == "return_call" {
                            "call"
                        } else {
                            "call_ref"
                        };
                        let call = map_instr_to_ast(inner.to_string(), args, span)?;
                        if let ExprKind::Call {
                            callee,
                            args: call_args,
                            ..
                        } = call.kind
                        {
                            let mut tail_args = vec![*callee];
                            tail_args.extend(call_args.into_iter().map(|a| a.value));
                            statements.push(Statement::with_span(
                                StmtKind::Expr(make_call("__wasm_return_call", tail_args, span)),
                                span,
                            ));
                        } else {
                            statements
                                .push(Statement::with_span(StmtKind::Return(Some(call)), span));
                        }
                    }
                    "br" => {
                        let target = br_target_of(args.first());
                        if let Some(entry) = labels.resolve(&target) {
                            // Unconditional branch: carry (consume) the top N stack
                            // values into the target, then jump. A `br` to a loop is
                            // a continue that carries the NEXT iteration's params; a
                            // `br` to a block carries the block's results.
                            let carry = branch_carry_temps(&entry);
                            carry_stack_into_temps(&carry, &mut stack, true, &mut statements);
                            statements.push(br_stmt_for(&entry, span));
                        } else {
                            statements.push(make_br_stmt_opt(None, labels, span));
                        }
                    }
                    "br_if" => {
                        // Arity 1 pops the condition; the label (if any) is the
                        // remaining immediate arg.
                        let mut lbl_arg: Option<&Expression> = None;
                        let mut cond: Option<Expression> = None;
                        if args.len() >= 2 {
                            lbl_arg = Some(&args[0]);
                            cond = Some(args[1].clone());
                        } else if args.len() == 1 {
                            cond = Some(args[0].clone());
                        }
                        let cond_expr = cond.unwrap_or(Expression::int(0));
                        let target = br_target_of(lbl_arg);
                        let mut then_body: Vec<Statement> = Vec::new();
                        let branch = match labels.resolve(&target) {
                            Some(entry) => {
                                // The carried values pass through a conditional
                                // branch, so peek (don't consume) the top N. A
                                // loop target carries its param temps; a block, its
                                // result temps.
                                let carry = branch_carry_temps(&entry);
                                carry_stack_into_temps(&carry, &mut stack, false, &mut then_body);
                                br_stmt_for(&entry, span)
                            }
                            None => make_br_stmt_opt(None, labels, span) };
                        then_body.push(branch);
                        statements.push(Statement::with_span(
                            StmtKind::If {
                                cond: cond_expr,
                                then_body,
                                else_body: None,
                                elifs: Vec::new() },
                            span,
                        ));
                    }
                    "br_table" => {
                        // `br_table l0 l1 … ln` pops a selector index and branches
                        // to the l_index frame (l_n is the default). Lower to an
                        // if/else-if chain over the index bound to a temp.
                        let targets: Vec<BrTarget> =
                            args.iter().map(|a| br_target_of(Some(a))).collect();
                        let index = stack.pop().unwrap_or(Expression::int(0));
                        // Exactly one target is taken, consuming the top N block
                        // results on the (post-selector) stack. Each case peeks the
                        // same snapshot and carries the top N into the chosen
                        // target's result temps before branching, mirroring `br`.
                        let carried: Vec<Expression> = stack.clone();
                        let br_for = |t: &BrTarget| -> Vec<Statement> {
                            match labels.resolve(t) {
                                Some(entry) => {
                                    let mut out = Vec::new();
                                    let n = entry.result_temps.len();
                                    let start = carried.len().saturating_sub(n);
                                    for (k, tmp) in entry.result_temps.iter().enumerate() {
                                        if let Some(val) = carried.get(start + k) {
                                            out.push(Statement::new(StmtKind::Expr(
                                                Expression::new(ExprKind::Assign {
                                                    target: Box::new(Expression::ident(tmp)),
                                                    value: Box::new(val.clone()) }),
                                            )));
                                        }
                                    }
                                    out.push(br_stmt_for(&entry, span));
                                    out
                                }
                                None => vec![make_br_stmt_opt(None, labels, span)] }
                        };
                        if targets.is_empty() {
                            // Degenerate: nothing to branch to.
                        } else if targets.len() == 1 {
                            statements.extend(br_for(&targets[0]));
                        } else {
                            let idx_tmp = fresh_result_temp();
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(idx_tmp.clone()),
                                    type_hint: None,
                                    init: Some(index),
                                    array_bounds: None,
                                    with_events: false }],
                                kind: VarDeclKind::Let }));
                            // Default (last) branch, then wrap each earlier case.
                            let mut chain = br_for(&targets[targets.len() - 1]);
                            for k in (0..targets.len() - 1).rev() {
                                let cond = Expression::new(ExprKind::Binary {
                                    op: BinOp::StrictEq,
                                    left: Box::new(Expression::ident(&idx_tmp)),
                                    right: Box::new(Expression::int(k as i64)) });
                                chain = vec![Statement::with_span(
                                    StmtKind::If {
                                        cond,
                                        then_body: br_for(&targets[k]),
                                        else_body: Some(chain),
                                        elifs: Vec::new() },
                                    span,
                                )];
                            }
                            statements.extend(chain);
                        }
                    }
                    // `br_on_cast L $from $to` branches to L (carrying the ref as
                    // the block result) when the ref IS `$to`; `br_on_cast_fail`
                    // when it is NOT. The ref stays on the stack for the
                    // fall-through path (like `br_if`'s peeked block result).
                    "br_on_cast" | "br_on_cast_fail" => {
                        let is_fail = name == "br_on_cast_fail";
                        let target = br_target_of(args.first());
                        let to_ht = args.get(2).cloned().unwrap_or(Expression::null());
                        // Consume the ref into a temp: on branch it becomes the
                        // block result; on fall-through it is pushed back so the
                        // continuation (e.g. `drop`) consumes it. Binding once
                        // avoids re-evaluating and keeps the stack balanced on
                        // both paths.
                        let ref_val = stack.pop().unwrap_or_else(Expression::null);
                        let tmp = fresh_result_temp();
                        statements.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(tmp.clone()),
                                type_hint: None,
                                init: Some(ref_val),
                                array_bounds: None,
                                with_events: false }],
                            kind: VarDeclKind::Let }));
                        let test =
                            make_call("ref_test", vec![to_ht, Expression::ident(&tmp)], span);
                        let cond = if is_fail {
                            Expression::new(ExprKind::Binary {
                                op: BinOp::StrictEq,
                                left: Box::new(test),
                                right: Box::new(Expression::int(0)) })
                        } else {
                            test
                        };
                        let mut then_body: Vec<Statement> = Vec::new();
                        match labels.resolve(&target) {
                            Some(entry) => {
                                // The cast ref is the target block's topmost result.
                                if let Some(rt) = entry.result_temps.last() {
                                    then_body.push(Statement::new(StmtKind::Expr(
                                        Expression::new(ExprKind::Assign {
                                            target: Box::new(Expression::ident(rt)),
                                            value: Box::new(Expression::ident(&tmp)) }),
                                    )));
                                }
                                then_body.push(br_stmt_for(&entry, span));
                            }
                            None => then_body.push(make_br_stmt_opt(None, labels, span)) }
                        statements.push(Statement::with_span(
                            StmtKind::If {
                                cond,
                                then_body,
                                else_body: None,
                                elifs: Vec::new() },
                            span,
                        ));
                        // Fall-through: the ref stays available for continuation.
                        stack.push(Expression::ident(&tmp));
                    }
                    _ => {
                        // A `call` yields as many values as the callee has results;
                        // a 0-result (void) call is a statement that must run in
                        // place, not a deferred stack value. Everything else that
                        // reaches here pushes a single value.
                        let pushes = if name == "call" {
                            call_result_count(&args)
                        } else {
                            get_instruction_push_count(&name)
                        };
                        let expr = map_instr_to_ast(name.clone(), args, span)?;
                        if name == "call" && pushes >= 2 {
                            // Multi-result call: bind the N stack results into N
                            // fresh temps via a destructure the shared compiler's
                            // multi-value ABI (`detect_multi_value_receive`) reads
                            // off the stack directly, then leave the temps on the
                            // value stack (temps[0] ← deepest) for downstream ops.
                            let temps: Vec<String> =
                                (0..pushes).map(|_| fresh_result_temp()).collect();
                            let pats: Vec<ArrayPatternElem> = temps
                                .iter()
                                .map(|t| {
                                    ArrayPatternElem::Pattern(
                                        BindingPattern::Ident(t.clone()),
                                        None,
                                    )
                                })
                                .collect();
                            statements.push(Statement::with_span(
                                StmtKind::Assign {
                                    targets: vec![Expression::new(ExprKind::Destructure(
                                        DestructurePattern::Array(pats),
                                    ))],
                                    value: expr,
                                    by_ref: false },
                                span,
                            ));
                            for t in &temps {
                                stack.push(Expression::ident(t));
                            }
                        } else if pushes > 0 {
                            stack.push(expr);
                        } else {
                            statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                        }
                    }
                }
            }
            _ => return Err(format!("Unexpected instr rule: {:?}", inner.as_rule())) }
    }

    // Flush remaining stack values as statements
    for expr in stack {
        statements.push(Statement::new(StmtKind::Expr(expr)));
    }

    Ok(statements)
}

fn get_instruction_arity(name: &str, args: &[Expression]) -> usize {
    // A `@@mem<N>` multi-memory selector suffix is not part of the op identity.
    let name = name.split_once("@@mem").map(|(b, _)| b).unwrap_or(name);
    match name {
        // Binary ops
        "i32.add" | "i32.sub" | "i32.mul" | "i32.div_s" | "i32.div_u" | "i32.rem_s"
        | "i32.rem_u" | "i32.and" | "i32.or" | "i32.xor" | "i32.shl" | "i32.shr_s"
        | "i32.shr_u" | "i32.rotl" | "i32.rotr" | "i64.add" | "i64.sub" | "i64.mul"
        | "i64.div_s" | "i64.div_u" | "i64.rem_s" | "i64.rem_u" | "i64.and" | "i64.or"
        | "i64.xor" | "i64.shl" | "i64.shr_s" | "i64.shr_u" | "i64.rotl" | "i64.rotr"
        | "f32.add" | "f32.sub" | "f32.mul" | "f32.div" | "f32.min" | "f32.max"
        | "f32.copysign" | "f64.add" | "f64.sub" | "f64.mul" | "f64.div" | "f64.min"
        | "f64.max" | "f64.copysign" | "i32.eq" | "i32.ne" | "i32.lt_s" | "i32.lt_u"
        | "i32.le_s" | "i32.le_u" | "i32.gt_s" | "i32.gt_u" | "i32.ge_s" | "i32.ge_u"
        | "i64.eq" | "i64.ne" | "i64.lt_s" | "i64.lt_u" | "i64.le_s" | "i64.le_u" | "i64.gt_s"
        | "i64.gt_u" | "i64.ge_s" | "i64.ge_u" | "f32.eq" | "f32.ne" | "f32.lt" | "f32.le"
        | "f32.gt" | "f32.ge" | "f64.eq" | "f64.ne" | "f64.lt" | "f64.le" | "f64.gt" | "f64.ge" => {
            2
        }

        // Unary / Conversion ops
        "i32.clz"
        | "i32.ctz"
        | "i32.popcnt"
        | "i32.eqz"
        | "i64.clz"
        | "i64.ctz"
        | "i64.popcnt"
        | "i64.eqz"
        | "f32.abs"
        | "f32.neg"
        | "f32.ceil"
        | "f32.floor"
        | "f32.trunc"
        | "f32.nearest"
        | "f32.sqrt"
        | "f64.abs"
        | "f64.neg"
        | "f64.ceil"
        | "f64.floor"
        | "f64.trunc"
        | "f64.nearest"
        | "f64.sqrt"
        | "i32.wrap_i64"
        | "i64.extend_i32_s"
        | "i64.extend_i32_u"
        | "i32.trunc_f32_s"
        | "i32.trunc_f32_u"
        | "i32.trunc_f64_s"
        | "i32.trunc_f64_u"
        | "i64.trunc_f32_s"
        | "i64.trunc_f32_u"
        | "i64.trunc_f64_s"
        | "i64.trunc_f64_u"
        | "f32.convert_i32_s"
        | "f32.convert_i32_u"
        | "f32.convert_i64_s"
        | "f32.convert_i64_u"
        | "f64.convert_i32_s"
        | "f64.convert_i32_u"
        | "f64.convert_i64_s"
        | "f64.convert_i64_u"
        | "f32.demote_f64"
        | "f64.promote_f32"
        | "i32.reinterpret_f32"
        | "i64.reinterpret_f64"
        | "f32.reinterpret_i32"
        | "f64.reinterpret_i64"
        | "i32.extend8_s"
        | "i32.extend16_s"
        | "i64.extend8_s"
        | "i64.extend16_s"
        | "i64.extend32_s"
        | "i32.trunc_sat_f32_s"
        | "i32.trunc_sat_f32_u"
        | "i32.trunc_sat_f64_s"
        | "i32.trunc_sat_f64_u"
        | "i64.trunc_sat_f32_s"
        | "i64.trunc_sat_f32_u"
        | "i64.trunc_sat_f64_s"
        | "i64.trunc_sat_f64_u" => 1,

        // Variable set / tee
        "local.set" | "global.set" | "local.tee" => 1,

        // Select
        "select" => 3,

        // Drop
        "drop" => 1,

        // Memory load/store
        "i32.load" | "i64.load" | "f32.load" | "f64.load" | "i32.load8_s" | "i32.load8_u"
        | "i32.load16_s" | "i32.load16_u" | "i64.load8_s" | "i64.load8_u" | "i64.load16_s"
        | "i64.load16_u" | "i64.load32_s" | "i64.load32_u" => 1, // address

        "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8" | "i32.store16"
        | "i64.store8" | "i64.store16" | "i64.store32" => 2, // address, value

        // Memory size / grow / bulk. fill/copy/init each pop 3 stack operands
        // (their data/mem-index selectors are immediates, not stack operands).
        "memory.size" => 0,
        "memory.grow" => 1,
        "memory.fill" | "memory.copy" | "memory.init" => 3,

        // Tables. The table index is an immediate; these are the stack operands.
        "table.get" => 1,                // elem index
        "table.set" | "table.grow" => 2, // (index,value) / (init,delta)
        "table.size" => 0,
        "table.fill" | "table.copy" | "table.init" => 3,

        // GC references without a type/field immediate — pure stack arity.
        "ref.i31" => 1,                 // i32 → i31ref
        "i31.get_s" | "i31.get_u" => 1, // i31ref → i32
        "ref.as_non_null" | "any.convert_extern" | "extern.convert_any" => 1,
        "ref.is_null" => 1, // [ref] → [i32]
        "ref.eq" => 2,
        // [ref] → [i32] (test) / [ref] → [ref] (cast). The heap-type operand is
        // an immediate, not a stack value; one ref is popped.
        "ref.test" | "ref.test_null" | "ref.cast" | "ref.cast_null" => 1,

        // ── Stringref proposal (stack-operand counts; $mem is an immediate) ──
        "string.new_utf8" | "string.new_wtf8" | "string.new_lossy_utf8" | "string.new_wtf16" => 2, // ptr, len (wtf16: ptr, codeunits)
        "string.new_utf8_array"
        | "string.new_wtf16_array"
        | "string.new_wtf8_array"
        | "string.new_lossy_utf8_array" => 3, // arr, start, end
        "string.measure_utf8" | "string.measure_wtf8" | "string.measure_wtf16" => 1,
        "string.encode_utf8"
        | "string.encode_wtf16"
        | "string.encode_lossy_utf8"
        | "string.encode_wtf8" => 2, // str, ptr
        "string.encode_utf8_array"
        | "string.encode_wtf16_array"
        | "string.encode_lossy_utf8_array"
        | "string.encode_wtf8_array" => 3, // str, arr, start
        "string.concat" | "string.eq" | "string.compare" => 2,
        "string.is_usv_sequence" | "string.as_wtf8" | "string.as_wtf16" | "string.as_iter" => 1,
        "stringview_iter.next" | "stringview_wtf16.length" => 1,
        // iterator advance/rewind/slice take (view, codepoints).
        "stringview_iter.advance" | "stringview_iter.rewind" | "stringview_iter.slice" => 2,
        // WTF-16 view: get_codeunit(view,pos)=2, slice(view,start,end)=3,
        // encode(view,ptr,pos,len)=4. WTF-8 view: advance(view,pos,bytes)=3,
        // slice(view,start,end)=3, encode_utf8(view,ptr,pos,bytes)=4.
        "stringview_wtf16.get_codeunit" => 2,
        "stringview_wtf16.slice" | "stringview_wtf8.advance" | "stringview_wtf8.slice" => 3,
        "stringview_wtf16.encode" | "stringview_wtf8.encode_utf8" => 4,
        "array.len" => 1, // arrayref → i32
        // Array ops carrying a type-index immediate (kept as an immediate arg):
        "array.new" => 2,         // value, length
        "array.new_default" => 1, // length
        // array.new_fixed $T N: typeidx + count are immediates; N stack values.
        "array.new_fixed" => args
            .get(1)
            .and_then(|a| {
                if let ExprKind::Lit(Literal::Int(n)) = &a.kind {
                    Some(*n as usize)
                } else {
                    None
                }
            })
            .unwrap_or(0),
        "array.get_s" | "array.get_u" => 2, // arrayref, index
        // Typeless array access (VM ignores the WAT typeidx → walker drops it):
        "array.get" => 2,  // arrayref, index
        "array.set" => 3,  // arrayref, index, value
        "array.fill" => 4, // arrayref, index, value, count
        "array.copy" => 5, // dst, dst_off, src, src_off, len (2 typeidxs dropped)
        // GC array-from-segment ops carry `typeidx` + `dataidx`/`elemidx` as
        // immediates; the stack operands are (offset, size) for new_* and
        // (array, dest_offset, src_offset, size) for init_*.
        "array.new_data" | "array.new_elem" => 2, // offset, size
        "array.init_data" | "array.init_elem" => 4, // array, dst_off, src_off, size

        // br_if
        "br_if" => 1,

        // Call
        // `return_call` is a tail call — same operand shape as `call`.
        "call" | "return_call" => {
            if let Some(first) = args.first() {
                match &first.kind {
                    ExprKind::Ident(n) => {
                        FUNC_NAME_ARITIES.with(|f| *f.borrow().get(n).unwrap_or(&1))
                    }
                    ExprKind::Lit(Literal::Int(idx)) => {
                        FUNC_INDEX_ARITIES.with(|f| *f.borrow().get(*idx as usize).unwrap_or(&1))
                    }
                    _ => 1 }
            } else {
                1
            }
        }

        "call_indirect" | "return_call_indirect" => 2,
        // call_ref pops the funcref plus the sig's params.
        "call_ref" | "return_call_ref" => {
            let params = args
                .first()
                .and_then(|e| match &e.kind {
                    ExprKind::Ident(n) => TYPE_FUNC_PARAMS.with(|m| m.borrow().get(n).copied()),
                    _ => None })
                .unwrap_or(0);
            1 + params
        }

        // GC struct ops
        // struct.new $T: typeidx is an immediate; field values come from stack.
        // We stored field counts by type name in STRUCT_FIELD_COUNTS.
        "struct.new" => {
            // args[0] is typeidx immediate (ident or int) — not a stack value.
            // Remaining stack operands = field count for that type.
            if let Some(first) = args.first() {
                let type_name = match &first.kind {
                    ExprKind::Ident(n) => n.clone(),
                    ExprKind::Lit(Literal::Int(i)) => i.to_string(),
                    _ => String::new() };
                STRUCT_FIELD_COUNTS.with(|f| *f.borrow().get(&type_name).unwrap_or(&0))
            } else {
                0
            }
        }
        "struct.new_default" => 0, // no stack operands; typeidx is immediate
        "struct.get" | "struct.get_s" | "struct.get_u" => 1, // pops 1 ref
        "struct.set" => 2,         // pops ref + val

        // ── SIMD v128: number of STACK operands (lane index / v128.const values
        //    are immediates, not stack operands). ────────────────────────────
        n if is_simd_instr(n) => simd_stack_arity(n),

        _ => 0 }
}

/// Is this a SIMD (v128) instruction mnemonic?
fn is_simd_instr(name: &str) -> bool {
    matches!(
        name.split_once('.').map(|(p, _)| p),
        Some("i8x16" | "i16x8" | "i32x4" | "i64x2" | "f32x4" | "f64x2" | "v128")
    )
}

/// How many STACK operands a SIMD op consumes (immediates excluded). Derived
/// from the op's shape, matching the VM's expectations.
fn simd_stack_arity(name: &str) -> usize {
    let op = name.split_once('.').map(|(_, o)| o).unwrap_or(name);
    if op == "const" {
        return 0;
    }
    if op.contains("replace_lane") {
        return 2; // vector + scalar (lane is immediate)
    }
    if op.contains("extract_lane") {
        return 1; // vector (lane is immediate)
    }
    if op.ends_with("splat") {
        return 1; // scalar
    }
    if op == "bitselect" || op.contains("relaxed_madd") || op.contains("relaxed_nmadd")
        || op.contains("laneselect")
        // `i32x4.relaxed_dot_i8x16_i7x16_add_s` takes a third accumulator vector
        // (the plain `i16x8.relaxed_dot_…_s` is a normal 2-operand op).
        || op.contains("relaxed_dot") && op.ends_with("add_s")
    {
        return 3;
    }
    if op.ends_with("_lane") {
        return 2; // load_lane / store_lane: address + vector (lane immediate)
    }
    if op.contains("load") {
        return 1; // v128.load, load*_splat, load*x*, load*_zero: address
    }
    if op.contains("store") {
        return 2; // address + vector
    }
    // Unary (single vector in → out).
    if op == "not"
        || op.ends_with("all_true")
        || op.ends_with("any_true")
        || op.ends_with("bitmask")
        || op == "abs"
        || op == "neg"
        || op == "sqrt"
        || op == "ceil"
        || op == "floor"
        || op == "nearest"
        || op == "popcnt"
        || op == "trunc"
        || op.starts_with("extend_")
        || op.starts_with("extadd_pairwise")
        || op.starts_with("convert")
        || op.starts_with("promote")
        || op.starts_with("demote")
        || op.starts_with("trunc_sat")
        || op.starts_with("relaxed_trunc")
    {
        return 1;
    }
    // Everything else is binary: add/sub/mul/div/min/max/logic/compare/shift/
    // avgr/narrow/extmul/dot/*_sat/pmin/pmax/q15mulr/shuffle/swizzle/relaxed_*.
    2
}

/// Result count of the function targeted by a `call`, from its first arg (the
/// callee id or index). Unknown callees default to 1 (assume value-producing) so
/// only functions we positively know to be void become statements.
fn call_result_count(args: &[Expression]) -> usize {
    match args.first().map(|e| &e.kind) {
        Some(ExprKind::Ident(n)) => FUNC_NAME_RESULTS
            .with(|f| f.borrow().get(n).copied())
            .unwrap_or(1),
        Some(ExprKind::Lit(Literal::Int(idx))) => FUNC_INDEX_RESULTS
            .with(|f| f.borrow().get(*idx as usize).copied())
            .unwrap_or(1),
        _ => 1 }
}

fn get_instruction_push_count(name: &str) -> usize {
    // A `@@mem<N>` multi-memory selector suffix is not part of the op identity.
    let name = name.split_once("@@mem").map(|(b, _)| b).unwrap_or(name);
    match name {
        // `unreachable` stays at 0 ON PURPOSE. It looks like it should push —
        // WASM makes it polymorphic, satisfying any result type — but a 0 sends
        // the folded form down the STATEMENT path, and `assign_last_n_exprs_to`
        // rewrites a branch's trailing `Expr` statement into
        // `__wat_res0 = <expr>`. That is what carries the trap into an
        // `if (result i32)` branch. Moving it to 1 pushed it on the value stack
        // instead, the branch body came out EMPTY, and the result temp kept its
        // null initialiser — measured.
        "local.set" | "global.set" | "drop" | "br_if" | "br" | "unreachable" | "nop"
        | "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8" | "i32.store16"
        | "i64.store8" | "i64.store16" | "i64.store32" | "struct.set"
        // Bulk-memory / table / segment ops yield NO value. Without this they
        // default to pushing 1 and get deferred to the block's stack flush,
        // running out of order (e.g. a `memory.fill` after the load that reads
        // it). `memory.grow`/`memory.size`/`table.grow`/`table.size`/`table.get`
        // DO produce a value and stay at the default 1.
        | "memory.fill" | "memory.copy" | "memory.init" | "data.drop"
        | "table.set" | "table.fill" | "table.copy" | "table.init" | "elem.drop"
        // GC array stores/copies write into an array and yield no value.
        | "array.set" | "array.copy" | "array.fill"
        | "array.init_data" | "array.init_elem"
        // SIMD stores also write memory and return nothing.
        | "v128.store" | "v128.store8_lane" | "v128.store16_lane"
        | "v128.store32_lane" | "v128.store64_lane" => 0,
        _ => 1 }
}
