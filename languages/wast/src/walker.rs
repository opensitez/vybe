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
use vybe_ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;
use std::collections::HashMap;

thread_local! {
    static FUNC_INDEX_ARITIES: RefCell<Vec<usize>> = RefCell::new(Vec::new());
    static FUNC_NAME_ARITIES: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // type name → number of fields (for struct.new arity)
    static STRUCT_FIELD_COUNTS: RefCell<HashMap<String, usize>> = RefCell::new(HashMap::new());
    // Module functions compile to static methods of this class; a `call $f` to a
    // defined function is reached as `ClassName.f(...)`.
    static MODULE_CLASS_NAME: RefCell<String> = RefCell::new(String::new());
    // Names of functions DEFINED in the module (not imports) — call targets that
    // must be qualified with the module class name. Imports resolve via the
    // profile builtin table by their local id, so they are excluded.
    static DEFINED_FUNC_NAMES: RefCell<std::collections::HashSet<String>> =
        RefCell::new(std::collections::HashSet::new());
}

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

    let mut index_arities = Vec::new();
    let mut name_arities = HashMap::new();

    // 1. Pre-scan imports
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::import_field {
                    let mut name: Option<String> = None;
                    let mut params_count = 0;
                    for sub in inner.into_inner() {
                        match sub.as_rule() {
                            Rule::id => name = Some(sub.as_str()[1..].to_string()),
                            Rule::param => {
                                let mut has_id = false;
                                let mut types_count = 0;
                                for p in sub.into_inner() {
                                    if p.as_rule() == Rule::id {
                                        has_id = true;
                                    } else if p.as_rule() == Rule::val_type {
                                        types_count += 1;
                                    }
                                }
                                params_count += if has_id { 1 } else { types_count };
                            }
                            _ => {}
                        }
                    }
                    index_arities.push(params_count);
                    if let Some(n) = name {
                        name_arities.insert(n, params_count);
                    }
                }
            }
        }
    }

    // 2. Pre-scan defined functions
    let mut defined_names: std::collections::HashSet<String> = std::collections::HashSet::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::func_field {
                    let mut name: Option<String> = None;
                    let mut params_count = 0;
                    for sub in inner.into_inner() {
                        match sub.as_rule() {
                            Rule::id => name = Some(sub.as_str()[1..].to_string()),
                            Rule::param => {
                                let mut has_id = false;
                                let mut types_count = 0;
                                for p in sub.into_inner() {
                                    if p.as_rule() == Rule::id {
                                        has_id = true;
                                    } else if p.as_rule() == Rule::val_type {
                                        types_count += 1;
                                    }
                                }
                                params_count += if has_id { 1 } else { types_count };
                            }
                            _ => {}
                        }
                    }
                    index_arities.push(params_count);
                    if let Some(n) = name {
                        defined_names.insert(n.clone());
                        name_arities.insert(n, params_count);
                    }
                }
            }
        }
    }

    FUNC_INDEX_ARITIES.with(|f| *f.borrow_mut() = index_arities);
    FUNC_NAME_ARITIES.with(|f| *f.borrow_mut() = name_arities);
    DEFINED_FUNC_NAMES.with(|f| *f.borrow_mut() = defined_names);

    // 3. Pre-scan struct type definitions to know field counts for struct.new arity
    let mut struct_counts: HashMap<String, usize> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::type_field {
                    let mut type_name: Option<String> = None;
                    let mut field_count = 0usize;
                    let mut is_struct = false;
                    for sub in inner.into_inner() {
                        match sub.as_rule() {
                            Rule::id => type_name = Some(sub.as_str()[1..].to_string()),
                            Rule::composite_type => {
                                if let Some(inner2) = sub.into_inner().next() {
                                    if inner2.as_rule() == Rule::struct_type {
                                        is_struct = true;
                                        field_count = inner2
                                            .into_inner()
                                            .filter(|p| p.as_rule() == Rule::field_def)
                                            .count();
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    if is_struct {
                        if let Some(name) = type_name {
                            struct_counts.insert(name, field_count);
                        }
                    }
                }
            }
        }
    }
    STRUCT_FIELD_COUNTS.with(|f| *f.borrow_mut() = struct_counts);

    // 4. Detect the WASI command entry. A module that exports a function as
    //    "_start" is a command module — instantiation runs `_start` with no
    //    driver. Explicit `(start $f)` fields are handled separately below; if
    //    one is present we don't also auto-run `_start`.
    let mut start_export_name: Option<String> = None;
    let mut has_start_field = false;
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        match inner.as_rule() {
            Rule::start_field => has_start_field = true,
            Rule::func_field => {
                let mut id: Option<String> = None;
                let mut exports_start = false;
                for sub in inner.into_inner() {
                    match sub.as_rule() {
                        Rule::id => id = Some(sub.as_str()[1..].to_string()),
                        Rule::export_inline => {
                            if let Some(s) =
                                sub.into_inner().find(|p| p.as_rule() == Rule::string)
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
    let class_name = name.clone();
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

    // Auto-run the command entry `_start` at instantiation (unless an explicit
    // `(start …)` field already scheduled a run).
    if !has_start_field {
        if let Some(entry) = start_export_name {
            // Functions are static methods of the module class, so the entry is
            // reached as `ModuleClass._start()`.
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class_name)),
                field: entry,
                null_safe: false,
            });
            result.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: Vec::new(),
                    optional: false,
                },
            ))));
        }
    }
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
                params = walk_typeuse_params(child)?;
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

    body.extend(fold_instructions(instr_pairs, &mut labels)?);

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
            let body = fold_instructions(instr_pairs, labels)?;
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
            let body = fold_instructions(instr_pairs, labels)?;
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
        let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
        for child in children {
            match child.as_rule() {
                Rule::id => label = Some(child.as_str()[1..].to_string()),
                Rule::instr => instr_pairs.push(child),
                _ => {}
            }
        }
        labels.push(label.clone(), kind.clone());
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

    // ── (try instr* (catch ...)*) ─────────────────────────────────────────
    if name == "try" {
        let mut args: Vec<Expression> = Vec::new();
        let mut instr_pairs = Vec::new();
        for child in children {
            if child.as_rule() == Rule::instr {
                instr_pairs.push(child);
            }
        }
        let body = fold_instructions(instr_pairs, labels)?;
        for stmt in body {
            if let StmtKind::Expr(e) = stmt.kind {
                args.push(e);
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
            Rule::catch_block | Rule::catch_all_block => {
                let mut instr_pairs = Vec::new();
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        instr_pairs.push(sub);
                    }
                }
                let body = fold_instructions(instr_pairs, labels)?;
                for stmt in body {
                    if let StmtKind::Expr(e) = stmt.kind {
                        args.push(e);
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
            // A call to a function DEFINED in this module targets a static method
            // of the module class; qualify `Ident(f)` as `ClassName.f`. Imports
            // keep their bare name so the profile builtin table resolves them.
            let callee = match &callee.kind {
                ExprKind::Ident(n)
                    if DEFINED_FUNC_NAMES.with(|d| d.borrow().contains(n)) =>
                {
                    let class = MODULE_CLASS_NAME.with(|c| c.borrow().clone());
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(Expression::ident(&class)),
                            field: n.clone(),
                            null_safe: false,
                        },
                        span,
                    )
                }
                _ => callee,
            };
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: call_args.into_iter().map(Argument::positional).collect(),
                    optional: false,
                },
                span,
            ))
        }

        // ── GC / WasmGC struct ops ────────────────────────────────────────
        // struct.new $T v0 v1 ...  → {"0": v0, "1": v1, ...}
        // args: [typeidx, field_val_0, field_val_1, ...]
        "struct.new" => {
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
                    value: v,
                })
                .collect();
            Ok(Expression::with_span(ExprKind::Object(props), span))
        }
        // struct.new_default $T → {}
        "struct.new_default" => Ok(Expression::with_span(ExprKind::Object(vec![]), span)),
        // struct.get $T N ref  → ref["N"]
        // args: [typeidx, fieldidx, ref_expr]
        "struct.get" | "struct.get_s" | "struct.get_u" => {
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
            Ok(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(obj),
                    field: field_idx.to_string(),
                    null_safe: false,
                },
                span,
            ))
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
                            null_safe: false,
                        },
                        span,
                    )),
                    value: Box::new(val),
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

fn fold_instructions(
    pairs: Vec<Pair<Rule>>,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    let mut stack: Vec<Expression> = Vec::new();
    let mut statements: Vec<Statement> = Vec::new();

    for pair in pairs {
        let span = to_span(&pair);
        let inner = if pair.as_rule() == Rule::instr {
            pair.into_inner().next().ok_or("Empty instr")?
        } else {
            pair
        };

        match inner.as_rule() {
            Rule::folded_instr => {
                let expr = walk_folded_instr_as_expr(inner, span, labels)?;
                stack.push(expr);
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

                // Parse inline arguments
                let mut args = Vec::new();
                for raw in raw_args {
                    args.push(walk_instr_arg_pair(raw, labels)?);
                }

                // Determine stack arity
                let arity = get_instruction_arity(&name, &args);
                if name == "call" {
                    eprintln!(
                        "DBG call args={:?} arity={} stack_len={}",
                        args.iter().map(|a| format!("{:?}", a.kind)).collect::<Vec<_>>(),
                        arity,
                        stack.len()
                    );
                }
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
                    "unreachable" => {
                        statements.push(Statement::with_span(
                            StmtKind::Throw {
                                expr: None,
                                cause: None,
                            },
                            span,
                        ));
                    }
                    "return" => {
                        let val = stack.pop();
                        statements.push(Statement::with_span(StmtKind::Return(val), span));
                    }
                    "br" => {
                        let lbl = args.first().and_then(|a| match &a.kind {
                            ExprKind::Ident(n) => Some(n.as_str()),
                            _ => None,
                        });
                        statements.push(make_br_stmt_opt(lbl, labels, span));
                    }
                    "br_if" => {
                        let mut lbl = None;
                        let mut cond = None;
                        if args.len() >= 2 {
                            if let ExprKind::Ident(ref n) = args[0].kind {
                                lbl = Some(n.clone());
                            }
                            cond = Some(args[1].clone());
                        } else if args.len() == 1 {
                            cond = Some(args[0].clone());
                        }
                        let cond_expr = cond.unwrap_or(Expression::int(0));
                        let branch = make_br_stmt_opt(lbl.as_deref(), labels, span);
                        statements.push(Statement::with_span(
                            StmtKind::If {
                                cond: cond_expr,
                                then_body: vec![branch],
                                else_body: None,
                                elifs: Vec::new(),
                            },
                            span,
                        ));
                    }
                    _ => {
                        // Value-producing or standard instruction.
                        let expr = map_instr_to_ast(name.clone(), args, span)?;
                        let pushes = get_instruction_push_count(&name);
                        if pushes > 0 {
                            stack.push(expr);
                        } else {
                            statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                        }
                    }
                }
            }
            _ => return Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
        }
    }

    // Flush remaining stack values as statements
    for expr in stack {
        statements.push(Statement::new(StmtKind::Expr(expr)));
    }

    Ok(statements)
}

fn get_instruction_arity(name: &str, args: &[Expression]) -> usize {
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

        // Memory size / grow
        "memory.size" => 0,
        "memory.grow" => 1,

        // br_if
        "br_if" => 1,

        // Call
        "call" => {
            if let Some(first) = args.first() {
                match &first.kind {
                    ExprKind::Ident(n) => {
                        FUNC_NAME_ARITIES.with(|f| *f.borrow().get(n).unwrap_or(&1))
                    }
                    ExprKind::Lit(Literal::Int(idx)) => {
                        FUNC_INDEX_ARITIES.with(|f| *f.borrow().get(*idx as usize).unwrap_or(&1))
                    }
                    _ => 1,
                }
            } else {
                1
            }
        }

        "call_indirect" => 2,

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
                    _ => String::new(),
                };
                STRUCT_FIELD_COUNTS.with(|f| *f.borrow().get(&type_name).unwrap_or(&0))
            } else {
                0
            }
        }
        "struct.new_default" => 0, // no stack operands; typeidx is immediate
        "struct.get" | "struct.get_s" | "struct.get_u" => 1, // pops 1 ref
        "struct.set" => 2,         // pops ref + val

        _ => 0,
    }
}

fn get_instruction_push_count(name: &str) -> usize {
    match name {
        "local.set" | "global.set" | "drop" | "br_if" | "br" | "unreachable" | "nop"
        | "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8" | "i32.store16"
        | "i64.store8" | "i64.store16" | "i64.store32" | "struct.set" => 0,
        _ => 1,
    }
}
