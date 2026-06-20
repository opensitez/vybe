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
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

// ── Label context ─────────────────────────────────────────────────────────────
// `br $label` targets a block (Break) or a loop (Continue).  We track which
// as we walk block/loop constructs.

#[derive(Clone, PartialEq)]
enum LabelKind {
    Block,
    Loop,
}

struct LabelStack(Vec<(Option<String>, LabelKind)>);

impl LabelStack {
    fn new() -> Self {
        LabelStack(Vec::new())
    }
    fn push(&mut self, name: Option<String>, kind: LabelKind) {
        self.0.push((name, kind));
    }
    fn pop(&mut self) {
        self.0.pop();
    }

    fn kind_of(&self, label: &str) -> Option<LabelKind> {
        for (name, kind) in self.0.iter().rev() {
            if name.as_deref() == Some(label) {
                return Some(kind.clone());
            }
        }
        None
    }
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
        Rule::module => {
            body.extend(walk_module(pair)?);
            Ok(())
        }
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
        _ => Ok(()),
    }
}

// ── Module ────────────────────────────────────────────────────────────────────

fn walk_module(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut module_name: Option<String> = None;
    let mut members: Vec<ClassMember> = Vec::new();
    let mut pre_stmts: Vec<Statement> = Vec::new(); // before class (globals)
    let mut post_stmts: Vec<Statement> = Vec::new(); // after class (start, exports, imports)

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
                        let (name, init) = walk_global_field(inner)?;
                        pre_stmts.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(name),
                                type_hint: None,
                                init: Some(init),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }));
                    }
                    Rule::start_field => {
                        post_stmts.push(Statement::new(StmtKind::Expr(walk_start_field(inner)?)));
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

    let mut result = pre_stmts;
    result.push(class);
    result.extend(post_stmts);
    Ok(result)
}

// ── Function field ────────────────────────────────────────────────────────────

fn walk_func_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut export_names: Vec<String> = Vec::new();
    let mut labels = LabelStack::new();

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
                params = walk_typeuse_params(child)?;
            }
            Rule::local => {
                body.extend(walk_local(child)?);
            }
            Rule::instr => {
                body.extend(walk_instr_as_stmts(child, &mut labels)?);
            }
            _ => {}
        }
    }

    if func_name.is_empty() {
        func_name = export_names
            .first()
            .cloned()
            .unwrap_or_else(|| "__wasm_func".to_string());
    }

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
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }
    if types.is_empty() {
        return Ok(Vec::new());
    }
    if let Some(n) = name {
        return Ok(vec![Param {
            name: n,
            type_hint: types.into_iter().next(),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }]);
    }
    Ok(types
        .into_iter()
        .enumerate()
        .map(|(i, t)| Param {
            name: format!("p{}", i),
            type_hint: Some(t),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        })
        .collect())
}

fn walk_local(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::val_type => types.push(child.as_str().to_string()),
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
                "f32" | "f64" => Expression::float(0.0),
                _ => Expression::null(),
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
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
    }
}

fn walk_instr_as_expr(pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(inner, span, labels),
        Rule::plain_instr => walk_plain_instr_as_expr(inner, span, labels),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
    }
}

// ── Plain instructions ────────────────────────────────────────────────────────

fn walk_plain_instr_as_stmts(
    pair: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    let mut name = String::new();
    let mut raw_args: Vec<Pair<Rule>> = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::instr_name => name = child.as_str().to_string(),
            Rule::instr_arg => raw_args.push(child),
            _ => {}
        }
    }

    match name.as_str() {
        // ── return ───────────────────────────────────────────────────────
        "return" => Ok(vec![Statement::with_span(StmtKind::Return(None), span)]),

        // ── unreachable = WASM trap ───────────────────────────────────────
        "unreachable" => Ok(vec![Statement::with_span(
            StmtKind::Throw {
                expr: None,
                cause: None,
            },
            span,
        )]),

        // ── br $label ─────────────────────────────────────────────────────
        "br" => {
            let lbl = raw_args
                .first()
                .and_then(|a| a.clone().into_inner().next())
                .filter(|c| c.as_rule() == Rule::id)
                .map(|c| c.as_str()[1..].to_string());
            Ok(vec![make_br_stmt_opt(lbl.as_deref(), labels, span)])
        }

        // ── br_if $label [cond] ───────────────────────────────────────────
        "br_if" => {
            let mut lbl: Option<String> = None;
            let mut cond: Option<Expression> = None;
            for raw in raw_args {
                let inner = raw.into_inner().next();
                if let Some(inner) = inner {
                    if inner.as_rule() == Rule::id && lbl.is_none() {
                        lbl = Some(inner.as_str()[1..].to_string());
                    } else {
                        cond = Some(instr_arg_inner_to_expr(inner));
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
                    elifs: Vec::new(),
                },
                span,
            )])
        }

        // ── everything else → expression statement ────────────────────────
        _ => {
            let mut args = Vec::new();
            for raw in raw_args {
                args.push(walk_instr_arg_pair(raw, labels)?);
            }
            let expr = map_instr_to_ast(name, args, span)?;
            Ok(vec![Statement::with_span(StmtKind::Expr(expr), span)])
        }
    }
}

fn walk_plain_instr_as_expr(
    pair: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Expression, String> {
    let mut name = String::new();
    let mut args: Vec<Expression> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::instr_name => name = child.as_str().to_string(),
            Rule::instr_arg => args.push(walk_instr_arg_pair(child, labels)?),
            _ => {}
        }
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
            labels.push(label.clone(), LabelKind::Block);
            let mut body: Vec<Statement> = Vec::new();
            for instr in instr_pairs {
                body.extend(walk_instr_as_stmts(instr, labels)?);
            }
            labels.pop();
            let block_stmt = Statement::with_span(StmtKind::Block(body), span);
            if let Some(lbl) = label {
                Ok(vec![Statement::with_span(
                    StmtKind::Labeled {
                        label: lbl,
                        body: Box::new(block_stmt),
                    },
                    span,
                )])
            } else {
                Ok(vec![block_stmt])
            }
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
            labels.push(label.clone(), LabelKind::Loop);
            let mut body: Vec<Statement> = Vec::new();
            for instr in instr_pairs {
                body.extend(walk_instr_as_stmts(instr, labels)?);
            }
            labels.pop();
            let while_stmt = Statement::with_span(
                StmtKind::While {
                    cond: Expression::bool(true),
                    body,
                    else_body: None,
                },
                span,
            );
            if let Some(lbl) = label {
                Ok(vec![Statement::with_span(
                    StmtKind::Labeled {
                        label: lbl,
                        body: Box::new(while_stmt),
                    },
                    span,
                )])
            } else {
                Ok(vec![while_stmt])
            }
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
        "unreachable" => Ok(vec![Statement::with_span(
            StmtKind::Throw {
                expr: None,
                cause: None,
            },
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
                    elifs: Vec::new(),
                },
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
    let mut all_children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if all_children.is_empty() {
        return Ok(Expression::null());
    }
    let name = if all_children[0].as_rule() == Rule::instr_name {
        all_children.remove(0).as_str().to_string()
    } else {
        String::new()
    };
    walk_folded_core(name, all_children, span, labels)
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
        let mut exprs: Vec<Expression> = Vec::new();
        for child in children {
            match child.as_rule() {
                Rule::id => label = Some(child.as_str()[1..].to_string()),
                Rule::instr => {
                    labels.push(label.clone(), kind.clone());
                    let e = walk_instr_as_expr(child, labels)?;
                    labels.pop();
                    exprs.push(e);
                    labels.push(label.clone(), kind.clone());
                }
                _ => {}
            }
        }
        // pop the extra pushes
        for _ in 0..exprs.len().saturating_sub(0) {
            labels.pop();
        }
        return Ok(exprs.into_iter().last().unwrap_or(Expression::null()));
    }

    // ── (try instr* (catch ...)*) ─────────────────────────────────────────
    if name == "try" {
        let mut args: Vec<Expression> = Vec::new();
        for child in children {
            if child.as_rule() == Rule::instr
                || child.as_rule() == Rule::catch_block
                || child.as_rule() == Rule::catch_all_block
            {
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        args.push(walk_instr_as_expr(sub, labels)?);
                    }
                }
            }
        }
        return Ok(make_call("__wasm_try", args, span));
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
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        then_exprs.push(walk_instr_as_expr(sub, labels)?);
                    }
                }
            }
            Rule::else_block => {
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        else_exprs.push(walk_instr_as_expr(sub, labels)?);
                    }
                }
            }
            Rule::catch_block | Rule::catch_all_block => {
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        args.push(walk_instr_as_expr(sub, labels)?);
                    }
                }
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
                else_: Box::new(else_val),
            },
            span,
        ));
    }

    map_instr_to_ast(name, args, span)
}

// ── map_instr_to_ast — WAT instruction name → common AST expression ───────────

fn map_instr_to_ast(name: String, args: Vec<Expression>, span: Span) -> Result<Expression, String> {
    match name.as_str() {
        // ── Constants ─────────────────────────────────────────────────────
        "i32.const" | "i64.const" => Ok(args.into_iter().next().unwrap_or(Expression::int(0))),
        "f32.const" | "f64.const" => Ok(args.into_iter().next().unwrap_or(Expression::float(0.0))),
        // wasm:js-string builtins — string.const "text" → string literal
        "string.const" => Ok(args.into_iter().next().unwrap_or(Expression::string(""))),

        // ── Local / global get → Ident ────────────────────────────────────
        "local.get" | "global.get" => {
            let idx = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(match &idx.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(format!("p{}", i)), span)
                }
                _ => idx,
            })
        }

        // ── Local / global set → Assign ───────────────────────────────────
        "local.set" | "global.set" => {
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target = match &target_raw.kind {
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(format!("p{}", i)), span)
                }
                _ => target_raw,
            };
            Ok(Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(target),
                    value: Box::new(value),
                },
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
                _ => "__tee_tmp".to_string(),
            };
            Ok(Expression::with_span(
                ExprKind::Sequence(vec![
                    Expression::with_span(
                        ExprKind::Assign {
                            target: Box::new(Expression::ident(&target_name)),
                            value: Box::new(value),
                        },
                        span,
                    ),
                    Expression::ident(&target_name),
                ]),
                span,
            ))
        }

        // ── Binary arithmetic ─────────────────────────────────────────────
        "i32.add" | "i64.add" | "f32.add" | "f64.add" => bin_op(args, BinOp::Add, span),
        "i32.sub" | "i64.sub" | "f32.sub" | "f64.sub" => bin_op(args, BinOp::Sub, span),
        "i32.mul" | "i64.mul" | "f32.mul" | "f64.mul" => bin_op(args, BinOp::Mul, span),
        "i32.div_s" | "i32.div_u" | "i64.div_s" | "i64.div_u" | "f32.div" | "f64.div" => {
            bin_op(args, BinOp::Div, span)
        }
        "i32.rem_s" | "i32.rem_u" | "i64.rem_s" | "i64.rem_u" => bin_op(args, BinOp::Mod, span),
        "i32.and" | "i64.and" => bin_op(args, BinOp::BitAnd, span),
        "i32.or" | "i64.or" => bin_op(args, BinOp::BitOr, span),
        "i32.xor" | "i64.xor" => bin_op(args, BinOp::BitXor, span),
        "i32.shl" | "i64.shl" => bin_op(args, BinOp::Shl, span),
        "i32.shr_s" | "i32.shr_u" | "i64.shr_s" | "i64.shr_u" => bin_op(args, BinOp::Shr, span),

        // ── Comparisons ───────────────────────────────────────────────────
        "i32.eq" | "i64.eq" | "f32.eq" | "f64.eq" => bin_op(args, BinOp::Eq, span),
        "i32.ne" | "i64.ne" | "f32.ne" | "f64.ne" => bin_op(args, BinOp::NotEq, span),
        "i32.lt_s" | "i32.lt_u" | "i64.lt_s" | "i64.lt_u" | "f32.lt" | "f64.lt" => {
            bin_op(args, BinOp::Lt, span)
        }
        "i32.gt_s" | "i32.gt_u" | "i64.gt_s" | "i64.gt_u" | "f32.gt" | "f64.gt" => {
            bin_op(args, BinOp::Gt, span)
        }
        "i32.le_s" | "i32.le_u" | "i64.le_s" | "i64.le_u" | "f32.le" | "f64.le" => {
            bin_op(args, BinOp::LtEq, span)
        }
        "i32.ge_s" | "i32.ge_u" | "i64.ge_s" | "i64.ge_u" | "f32.ge" | "f64.ge" => {
            bin_op(args, BinOp::GtEq, span)
        }

        // ── eqz → == 0 ───────────────────────────────────────────────────
        "i32.eqz" | "i64.eqz" => {
            let operand = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(operand),
                    right: Box::new(Expression::int(0)),
                },
                span,
            ))
        }

        // ── Unary negation ────────────────────────────────────────────────
        "f32.neg" | "f64.neg" => {
            let operand = args.into_iter().next().unwrap_or(Expression::float(0.0));
            Ok(Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(operand),
                },
                span,
            ))
        }

        // ── select → ternary ──────────────────────────────────────────────
        "select" => {
            let mut it = args.into_iter();
            let val1 = it.next().unwrap_or(Expression::null());
            let val2 = it.next().unwrap_or(Expression::null());
            let cond = it.next().unwrap_or(Expression::bool(false));
            Ok(Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(cond),
                    then: Box::new(val1),
                    else_: Box::new(val2),
                },
                span,
            ))
        }

        // ── drop → evaluate and discard ───────────────────────────────────
        "drop" => Ok(args.into_iter().next().unwrap_or(Expression::null())),

        // ── nop ───────────────────────────────────────────────────────────
        "nop" => Ok(Expression::with_span(ExprKind::Lit(Literal::Null), span)),

        // ── unreachable / return / br in expression context ───────────────
        // These are meaningful at statement level; here they produce null.
        "unreachable" | "return" | "br" | "br_if" | "br_table" => Ok(Expression::null()),

        // ── call → Call(callee, args) ─────────────────────────────────────
        "call" => {
            let mut it = args.into_iter();
            let callee = it.next().unwrap_or(Expression::null());
            let call_args: Vec<Expression> = it.collect();
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: call_args.into_iter().map(Argument::positional).collect(),
                    optional: false,
                },
                span,
            ))
        }

        // ── everything else → call with dots replaced by underscores ──────
        _ => Ok(make_call(&name.replace('.', "_"), args, span)),
    }
}

// ── Instruction argument helpers ──────────────────────────────────────────────

fn walk_instr_arg_pair(pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty instr_arg")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(inner, Span::default(), labels),
        _ => Ok(instr_arg_inner_to_expr(inner)),
    }
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
        | Rule::mem_arg
        | Rule::val_lane_type => Expression::string(inner.as_str()),
        _ => Expression::null(),
    }
}

fn walk_index(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty index")?;
    match inner.as_rule() {
        Rule::integer => Ok(parse_integer(inner.as_str())),
        Rule::id => Ok(Expression::ident(&inner.as_str()[1..])),
        _ => Ok(Expression::null()),
    }
}

// ── Break/continue helper ─────────────────────────────────────────────────────

fn make_br_stmt_opt(label: Option<&str>, labels: &LabelStack, span: Span) -> Statement {
    match label {
        Some(lbl) => match labels.kind_of(lbl) {
            Some(LabelKind::Loop) => Statement::with_span(
                StmtKind::Continue(ContinueTarget::Label(lbl.to_string())),
                span,
            ),
            _ => Statement::with_span(StmtKind::Break(BreakTarget::Label(lbl.to_string())), span),
        },
        None => Statement::with_span(StmtKind::Break(BreakTarget::Implicit), span),
    }
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

fn walk_global_field(pair: Pair<Rule>) -> Result<(String, Expression), String> {
    let mut name = "__wasm_global".to_string();
    let mut init = Expression::int(0);
    let mut labels = LabelStack::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = child.as_str()[1..].to_string(),
            Rule::instr => init = walk_instr_as_expr(child, &mut labels)?,
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
    Ok(Statement::with_span(
        StmtKind::Expr(Expression::with_span(
            ExprKind::Call {
                callee: Box::new(Expression::ident(&func_name)),
                args: args.into_iter().map(Argument::positional).collect(),
                optional: false,
            },
            span,
        )),
        span,
    ))
}

fn walk_assert_return(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    let mut expected: Vec<Expression> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action => action_expr = Some(walk_action(child)?),
            Rule::result_val => expected.push(walk_const_expr(child)?),
            _ => {}
        }
    }
    let mut args = Vec::new();
    if let Some(a) = action_expr {
        args.push(a);
    }
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
            optional: false,
        },
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
            right: Box::new(right),
        },
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
        return Expression::int(s.parse::<i64>().unwrap_or(0));
    };
    let v = i64::from_str_radix(digits, 16).unwrap_or(0);
    Expression::int(if neg { -v } else { v })
}

fn parse_float(s: &str) -> Expression {
    let s = s.trim();
    match s {
        "inf" | "+inf" => Expression::float(f64::INFINITY),
        "-inf" => Expression::float(f64::NEG_INFINITY),
        "nan" | "+nan" | "-nan" => Expression::float(f64::NAN),
        _ if s.contains("nan:0x") => Expression::float(f64::NAN),
        _ => Expression::float(s.parse::<f64>().unwrap_or(0.0)),
    }
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
        end_col: end.1 as u32,
    }
}
