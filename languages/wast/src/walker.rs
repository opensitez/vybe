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
    // Exported function name → the static-method name it maps to on the module
    // class, so a WAST script `(invoke "name" …)` resolves to `Class.method`.
    static EXPORT_FUNC_MAP: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    // Monotonic counter for synthetic result temporaries of value-producing
    // structured control (block/if with a `(result …)` type).
    static WAST_TEMP_COUNTER: RefCell<usize> = const { RefCell::new(0) };
    // Exception tag name (without `$`) → payload arity, from `(tag $e (param …))`.
    // A `catch $e` needs the arity to bind the right number of payload values.
    static TAG_ARITIES: RefCell<HashMap<String, u8>> = RefCell::new(HashMap::new());
    // Stack of the currently-open catch handlers' captured `exnref` locals, so a
    // `rethrow` inside a catch body resolves to the exception it caught.
    static ACTIVE_CATCH_EXNREFS: RefCell<Vec<String>> = RefCell::new(Vec::new());
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
            None => return false,
        }
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
            None => return 0,
        }
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

/// Does an unfolded `block`/`loop`/`if` opener carry a `block_type` immediate
/// (`(result …)`) — i.e. does it produce a value on the stack?
fn peek_has_block_type(pair: &Pair<Rule>) -> bool {
    let inner = if pair.as_rule() == Rule::instr {
        match pair.clone().into_inner().next() {
            Some(p) => p,
            None => return false,
        }
    } else {
        pair.clone()
    };
    if inner.as_rule() != Rule::plain_instr {
        return false;
    }
    inner.into_inner().any(|c| {
        c.as_rule() == Rule::instr_arg
            && c.into_inner().next().map(|i| i.as_rule()) == Some(Rule::block_type)
    })
}

/// Rewrite the final value-producing statement of a branch body into an
/// assignment to `tmp`, so the branch's stack result is captured.
fn assign_last_expr_to(body: &mut [Statement], tmp: &str) {
    if let Some(last) = body.last_mut() {
        if let StmtKind::Expr(e) = &last.kind {
            let value = e.clone();
            last.kind = StmtKind::Expr(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident(tmp)),
                value: Box::new(value),
            }));
        }
    }
}

// ── Label context ─────────────────────────────────────────────────────────────
// `br $label` targets a block (Break) or a loop (Continue).  We track which
// as we walk block/loop constructs.

#[derive(Clone, PartialEq)]
enum LabelKind {
    Block,
    Loop,
}

#[derive(Clone)]
struct LabelEntry {
    /// Always present — a synthetic name is minted when the source omits one, so
    /// every block/loop is addressable (numeric `br N` needs no source label).
    name: String,
    kind: LabelKind,
    /// The result temporary for a value-producing block/loop; `br` to this frame
    /// carries the top of stack into it.
    result_temp: Option<String>,
}

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
    fn push(&mut self, name: Option<String>, kind: LabelKind, result_temp: Option<String>) -> String {
        let effective = name.unwrap_or_else(fresh_block_label);
        self.0.push(LabelEntry {
            name: effective.clone(),
            kind,
            result_temp,
        });
        effective
    }
    fn pop(&mut self) {
        self.0.pop();
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
            BrTarget::Innermost => self.0.last().cloned(),
        }
    }
}

/// How a `br`/`br_if` names its destination frame.
enum BrTarget {
    Named(String),
    Index(usize),
    Innermost,
}

/// Derive a `br` target from its first argument (label id or numeric index).
fn br_target_of(arg: Option<&Expression>) -> BrTarget {
    match arg.map(|a| &a.kind) {
        Some(ExprKind::Ident(n)) => BrTarget::Named(n.clone()),
        Some(ExprKind::Lit(Literal::Int(i))) => BrTarget::Index(*i as usize),
        _ => BrTarget::Innermost,
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
        ),
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

    // 1. Pre-scan imports. Params live inside `typeuse` (and, for imports, inside
    //    `import_desc`), so the signature scan must descend, not read direct children.
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::import_field {
                    let (name, params_count) = scan_func_signature(inner);
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
    let mut export_map: HashMap<String, String> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::func_field {
                    let (name, params_count) = scan_func_signature(inner.clone());
                    index_arities.push(params_count);
                    if let Some(n) = &name {
                        defined_names.insert(n.clone());
                        name_arities.insert(n.clone(), params_count);
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
    DEFINED_FUNC_NAMES.with(|f| *f.borrow_mut() = defined_names);
    EXPORT_FUNC_MAP.with(|f| *f.borrow_mut() = export_map);

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
                    // Linear memory + data segments: emitted before the class so
                    // the compiler lowers them into the script chunk's memory /
                    // data tables (the VM allocates pages and writes active data
                    // at instantiation, before `_start`).
                    Rule::memory_field => pre_stmts.push(walk_memory_field(inner)?),
                    Rule::data_field => pre_stmts.push(walk_data_field(inner)?),
                    Rule::table_field => pre_stmts.push(walk_table_field(inner)?),
                    // Exception tags: declared before the class so the tag
                    // entity exists in the script chunk; `throw`/`catch` in the
                    // function chunks re-import by name and coalesce to it.
                    Rule::tag_field => pre_stmts.push(walk_tag_field(inner)?),
                    _ => {} // elem, type — structural metadata
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

/// Recursively read a func/import field's signature: its (first) id and its
/// parameter count. Parameters are wrapped in `typeuse`, and imported funcs are
/// further wrapped in `import_desc`, so a flat scan of direct children misses
/// them — the call-site arity would then be 0 and stack operands never consumed.
fn scan_func_signature(pair: Pair<Rule>) -> (Option<String>, usize) {
    let mut name: Option<String> = None;
    let mut count = 0usize;
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
            Rule::typeuse | Rule::import_desc => {
                let (n, c) = scan_func_signature(child);
                if name.is_none() {
                    name = n;
                }
                count += c;
            }
            _ => {}
        }
    }
    (name, count)
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
            let effective = labels.push(label.clone(), LabelKind::Block, None);
            let body = fold_instructions(instr_pairs, labels)?;
            labels.pop();
            let block_stmt = Statement::with_span(StmtKind::Block(body), span);
            Ok(vec![Statement::with_span(
                StmtKind::Labeled {
                    label: effective,
                    body: Box::new(block_stmt),
                },
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
            let effective = labels.push(label.clone(), LabelKind::Loop, None);
            let mut body = fold_instructions(instr_pairs, labels)?;
            labels.pop();
            // A WASM loop exits on fall-through; while(true) needs an explicit break.
            body.push(Statement::with_span(StmtKind::Break(BreakTarget::Implicit), span));
            let while_stmt = Statement::with_span(
                StmtKind::While {
                    cond: Expression::bool(true),
                    body,
                    else_body: None,
                },
                span,
            );
            Ok(vec![Statement::with_span(
                StmtKind::Labeled {
                    label: effective,
                    body: Box::new(while_stmt),
                },
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
        labels.push(label.clone(), kind.clone(), None);
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
            right: Box::new(r),
        })
    }
    fn tern(c: Expression, t: Expression, e: Expression) -> Expression {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(c),
            then: Box::new(t),
            else_: Box::new(e),
        })
    }
    let zero = || Expression::float(0.0);
    // whole number → "<x>.0"; otherwise the natural decimal string.
    let finite = tern(
        bin(BinOp::StrictEq, bin(BinOp::Mod, x.clone(), Expression::float(1.0)), zero()),
        bin(
            BinOp::Add,
            bin(BinOp::Add, x.clone(), Expression::string("")),
            Expression::string(".0"),
        ),
        bin(BinOp::Add, x.clone(), Expression::string("")),
    );
    // (x - x) is 0 for finite values but NaN for ±∞.
    let inf_or_finite = tern(
        bin(BinOp::StrictNotEq, bin(BinOp::Sub, x.clone(), x.clone()), zero()),
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

fn map_instr_to_ast(name: String, args: Vec<Expression>, span: Span) -> Result<Expression, String> {
    match name.as_str() {
        // Typeless array access: the WAT typeidx (`$t`) immediates are the first
        // arg(s) but the VM's array.get/set/fill/copy don't read them — drop and
        // keep only the stack operands. array.copy carries two typeidxs.
        "array.get" | "array.set" | "array.fill" => {
            let rest: Vec<Expression> = args.into_iter().skip(1).collect();
            Ok(make_call(&name.replace('.', "_"), rest, span))
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
                Ok(Expression::with_span(ExprKind::Lit(Literal::BigInt(*n)), span))
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
                    expr: Box::new(operand),
                },
                span,
            ))
        }

        // ── select → ternary ──────────────────────────────────────────────
        "select" => {
            // `select (result t)` prepends a result-type annotation; the stack
            // operands (val1, val2, cond) are always the last three args.
            let n = args.len();
            let val1 = args.get(n.wrapping_sub(3)).cloned().unwrap_or(Expression::null());
            let val2 = args.get(n.wrapping_sub(2)).cloned().unwrap_or(Expression::null());
            let cond = args.get(n.wrapping_sub(1)).cloned().unwrap_or(Expression::bool(false));
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

        // ── GC / WasmGC reference ops ─────────────────────────────────────
        // ref.null <heaptype> pushes a typed null reference. The heap type is
        // an immediate annotation, not a stack value, and the VM has a single
        // null — so drop the arg and produce a plain null (like `nop`). Applies
        // to bare heap types (`func`/`extern`) and indexed types (`$T`) alike.
        "ref.null" => Ok(Expression::with_span(ExprKind::Lit(Literal::Null), span)),

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
        // array.new_fixed $T N v0 v1 … → [v0, v1, …]. args: [typeidx, N, v0…].
        // The typeidx + count immediates are dropped (VM arrays are typeless);
        // the N popped stack values become an array literal.
        "array.new_fixed" => {
            let vals: Vec<ArrayElement> = if args.len() > 2 {
                args[2..]
                    .iter()
                    .map(|v| ArrayElement {
                        key: None,
                        value: v.clone(),
                        spread: false,
                        by_ref: false,
                    })
                    .collect()
            } else {
                vec![]
            };
            Ok(Expression::with_span(ExprKind::Array(vals), span))
        }
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
        | Rule::bare_heap_type
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
            max_pages,
        },
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
            arity,
        },
        span,
    ))
}

fn walk_table_field(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut min_size: u64 = 0;
    let mut max_size: Option<u64> = None;
    for child in pair.into_inner() {
        if child.as_rule() == Rule::table_type {
            // table_type = integer integer? ref_type — pages then optional max.
            let mut nums = child.into_inner().filter(|p| p.as_rule() == Rule::integer);
            if let Some(min) = nums.next() {
                min_size = parse_wat_u64(min.as_str());
            }
            if let Some(max) = nums.next() {
                max_size = Some(parse_wat_u64(max.as_str()));
            }
        }
    }
    Ok(Statement::with_span(
        StmtKind::TableDecl { min_size, max_size },
        span,
    ))
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
            bytes,
        },
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
            _ => 0,
        }
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
            c if c.is_ascii_hexdigit() && i + 2 < bytes.len() && bytes[i + 2].is_ascii_hexdigit() => {
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
                    null_safe: false,
                },
                span,
            )
        }
        None => Expression::ident(&func_name),
    };
    Ok(Statement::with_span(
        StmtKind::Expr(Expression::with_span(
            ExprKind::Call {
                callee: Box::new(callee),
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
        None => (false, s.strip_prefix('+').unwrap_or(s)),
    };
    let rest = rest.strip_prefix("0x").or_else(|| rest.strip_prefix("0X"))?;
    let (mantissa, exp) = match rest.find(['p', 'P']) {
        Some(i) => (&rest[..i], rest[i + 1..].parse::<i32>().ok()?),
        None => (rest, 0),
    };
    let (int_part, frac_part) = match mantissa.find('.') {
        Some(i) => (&mantissa[..i], &mantissa[i + 1..]),
        None => (mantissa, ""),
    };
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
        end_col: end.1 as u32,
    }
}

/// Peek the plain-instruction keyword of an `instr`/`plain_instr` pair without
/// consuming it. Returns None for folded instructions (which carry no linear
/// `block`/`loop`/`if`/`else`/`end` tokens).
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

/// How an unfolded `try` block terminates: a structural `end`, or a legacy
/// `delegate N` (which reraises to an enclosing try and has no `end`).
enum TryTerminator {
    End(usize),
    Delegate(usize),
}

/// Scan an unfolded `try … (catch $e … | catch_all …)* (end | delegate N)`,
/// returning its top-level catch clauses (`(tag, keyword_index)`, `tag == None`
/// for `catch_all`) and the terminator index. Nesting is tracked with an
/// opener stack so a nested `try … delegate` (which has no `end`) closes its
/// own level correctly.
fn scan_try(
    pairs: &[Pair<Rule>],
    opener: usize,
) -> Result<(Vec<(Option<String>, usize)>, TryTerminator), String> {
    // true = the open level is a `try`; false = block/loop/if.
    let mut stack: Vec<bool> = vec![true];
    let mut clauses: Vec<(Option<String>, usize)> = Vec::new();
    let mut j = opener + 1;
    while j < pairs.len() {
        if let Some(kw) = peek_plain_name(&pairs[j]) {
            match kw.as_str() {
                "block" | "loop" | "if" => stack.push(false),
                "try" => stack.push(true),
                "catch" if stack.len() == 1 => {
                    clauses.push((peek_plain_label(&pairs[j]), j));
                }
                "catch_all" if stack.len() == 1 => clauses.push((None, j)),
                "delegate" => {
                    if *stack.last().unwrap_or(&false) {
                        stack.pop();
                        if stack.is_empty() {
                            return Ok((clauses, TryTerminator::Delegate(j)));
                        }
                    }
                }
                "end" => {
                    stack.pop();
                    if stack.is_empty() {
                        return Ok((clauses, TryTerminator::End(j)));
                    }
                }
                _ => {}
            }
        }
        j += 1;
    }
    Err("unterminated try (missing end/delegate)".to_string())
}

/// Does this flat pair slice contain a top-level `rethrow`? A catch whose body
/// rethrows must capture the exception's `exnref` (catch_ref), so the reraise
/// lowers to `throw_ref`.
fn pairs_contain_rethrow(pairs: &[Pair<Rule>]) -> bool {
    pairs
        .iter()
        .any(|p| peek_plain_name(p).as_deref() == Some("rethrow"))
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
                        let produces_value = peek_has_block_type(&pairs[i]);
                        // Condition is the value on top of the stack.
                        let cond = stack.pop().unwrap_or(Expression::bool(false));
                        for e in stack.drain(..) {
                            statements.push(Statement::new(StmtKind::Expr(e)));
                        }
                        let then_end = else_idx.unwrap_or(end_idx);
                        let then_pairs: Vec<Pair<Rule>> = pairs[i + 1..then_end].to_vec();
                        labels.push(label.clone(), LabelKind::Block, None);
                        let mut then_body = fold_instructions(then_pairs, labels)?;
                        let mut else_body = if let Some(ei) = else_idx {
                            let else_pairs: Vec<Pair<Rule>> = pairs[ei + 1..end_idx].to_vec();
                            Some(fold_instructions(else_pairs, labels)?)
                        } else {
                            None
                        };
                        labels.pop();
                        // A `(result …)` if yields a value: capture each branch's
                        // trailing value in a temp and leave that temp on the stack.
                        if produces_value {
                            let tmp = fresh_result_temp();
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(tmp.clone()),
                                    type_hint: None,
                                    init: Some(Expression::null()),
                                    array_bounds: None,
                                    with_events: false,
                                }],
                                kind: VarDeclKind::Let,
                            }));
                            assign_last_expr_to(&mut then_body, &tmp);
                            if let Some(eb) = else_body.as_mut() {
                                assign_last_expr_to(eb, &tmp);
                            }
                            statements.push(Statement::with_span(
                                StmtKind::If {
                                    cond,
                                    then_body,
                                    else_body,
                                    elifs: Vec::new(),
                                },
                                span,
                            ));
                            stack.push(Expression::ident(&tmp));
                        } else {
                            statements.push(Statement::with_span(
                                StmtKind::If {
                                    cond,
                                    then_body,
                                    else_body,
                                    elifs: Vec::new(),
                                },
                                span,
                            ));
                        }
                    } else {
                        // block / loop take no condition. Pop the block's
                        // param values off the top to seed the body, then
                        // sequence any remaining pending side effects.
                        let param_count = peek_block_param_count(&pairs[i]);
                        let seed = if param_count > 0 && stack.len() >= param_count {
                            stack.split_off(stack.len() - param_count)
                        } else {
                            Vec::new()
                        };
                        for e in stack.drain(..) {
                            statements.push(Statement::new(StmtKind::Expr(e)));
                        }
                        let body_pairs: Vec<Pair<Rule>> = pairs[i + 1..end_idx].to_vec();
                        // Loop parameters thread values across iterations, which
                        // this lowering can't model — treat such loops as a
                        // one-shot block so `br` breaks (terminates) rather than
                        // continues forever.
                        let loop_has_param =
                            kw == "loop" && peek_opener_has_param(&pairs[i]);
                        let kind = if kw == "block" || loop_has_param {
                            LabelKind::Block
                        } else {
                            LabelKind::Loop
                        };
                        // A `(result …)` block/loop yields a value: `br` to it
                        // carries the stack top into a temp, and the fall-through
                        // value assigns the same temp; the temp is left on the stack.
                        let result_temp = if peek_has_block_type(&pairs[i]) {
                            Some(fresh_result_temp())
                        } else {
                            None
                        };
                        if let Some(tmp) = &result_temp {
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(tmp.clone()),
                                    type_hint: None,
                                    init: Some(Expression::null()),
                                    array_bounds: None,
                                    with_events: false,
                                }],
                                kind: VarDeclKind::Let,
                            }));
                        }
                        let effective =
                            labels.push(label.clone(), kind, result_temp.clone());
                        let mut body = fold_instructions_seeded(body_pairs, labels, seed)?;
                        labels.pop();
                        // Capture the fall-through value (unreachable if the body
                        // always branches out, which is why it's safe to append).
                        if let Some(tmp) = &result_temp {
                            assign_last_expr_to(&mut body, tmp);
                        }
                        let inner_stmt = if kw == "block" || loop_has_param {
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
                                    else_body: None,
                                },
                                span,
                            )
                        };
                        statements.push(Statement::with_span(
                            StmtKind::Labeled {
                                label: effective,
                                body: Box::new(inner_stmt),
                            },
                            span,
                        ));
                        if let Some(tmp) = &result_temp {
                            stack.push(Expression::ident(tmp));
                        }
                    }
                    i = end_idx + 1;
                    continue;
                }
                // ── Unfolded exception handling: try … catch … end ──────────
                "try" => {
                    let span = to_span(&pairs[i]);
                    let produces_value = peek_has_block_type(&pairs[i]);
                    let (clauses, terminator) = scan_try(&pairs, i)?;
                    let end_idx = match &terminator {
                        TryTerminator::End(e) => *e,
                        TryTerminator::Delegate(d) => *d,
                    };

                    // Sequence any pending side effects before the try.
                    for e in stack.drain(..) {
                        statements.push(Statement::new(StmtKind::Expr(e)));
                    }

                    // A `try (result T)` yields a value: capture the body's and
                    // each handler's trailing value in a shared temp, left on the
                    // stack afterwards (mirrors the block/loop lowering).
                    let result_temp = if produces_value {
                        Some(fresh_result_temp())
                    } else {
                        None
                    };
                    if let Some(tmp) = &result_temp {
                        statements.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(tmp.clone()),
                                type_hint: None,
                                init: Some(Expression::null()),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }));
                    }

                    // Body: opener+1 .. first clause (or the terminator).
                    let body_end = clauses.first().map(|(_, idx)| *idx).unwrap_or(end_idx);
                    let body_pairs: Vec<Pair<Rule>> = pairs[i + 1..body_end].to_vec();
                    let mut body = fold_instructions(body_pairs, labels)?;
                    if let Some(tmp) = &result_temp {
                        assign_last_expr_to(&mut body, tmp);
                    }

                    // Catch clauses. The delivered payload binds to fresh locals,
                    // seeded into the handler fold so its body reads them.
                    let mut wasm_catches: Vec<WasmCatch> = Vec::new();
                    for (k, (tag, kw_idx)) in clauses.iter().enumerate() {
                        let clause_body_start = kw_idx + 1;
                        let clause_body_end =
                            clauses.get(k + 1).map(|(_, idx)| *idx).unwrap_or(end_idx);
                        let clause_pairs: Vec<Pair<Rule>> =
                            pairs[clause_body_start..clause_body_end].to_vec();

                        let arity = tag.as_deref().map(tag_arity).unwrap_or(0);
                        let payload_binds: Vec<String> =
                            (0..arity).map(|_| fresh_result_temp()).collect();
                        let seed: Vec<Expression> =
                            payload_binds.iter().map(|n| Expression::ident(n)).collect();

                        // Only capture the exnref when the handler rethrows.
                        let capture_ref = pairs_contain_rethrow(&clause_pairs);
                        let exnref_bind = if capture_ref {
                            Some(fresh_result_temp())
                        } else {
                            None
                        };
                        if let Some(exnref) = &exnref_bind {
                            ACTIVE_CATCH_EXNREFS
                                .with(|s| s.borrow_mut().push(exnref.clone()));
                        }
                        let mut cbody = fold_instructions_seeded(clause_pairs, labels, seed)?;
                        if exnref_bind.is_some() {
                            ACTIVE_CATCH_EXNREFS.with(|s| {
                                s.borrow_mut().pop();
                            });
                        }
                        if let Some(tmp) = &result_temp {
                            assign_last_expr_to(&mut cbody, tmp);
                        }
                        wasm_catches.push(WasmCatch {
                            tag: tag.clone(),
                            payload_binds,
                            capture_ref,
                            exnref_bind,
                            body: cbody,
                        });
                    }

                    // Legacy `delegate N`: no catch clause — reraise to the
                    // enclosing try. Modelled as a catch_all_ref handler that
                    // `throw_ref`s the captured exnref (propagates outward).
                    if matches!(terminator, TryTerminator::Delegate(_)) {
                        let exnref = fresh_result_temp();
                        wasm_catches.push(WasmCatch {
                            tag: None,
                            payload_binds: Vec::new(),
                            capture_ref: true,
                            exnref_bind: Some(exnref.clone()),
                            body: vec![Statement::new(StmtKind::WasmRethrow {
                                exnref_local: exnref,
                            })],
                        });
                    }

                    statements.push(Statement::with_span(
                        StmtKind::WasmTryTable {
                            body,
                            catches: wasm_catches,
                        },
                        span,
                    ));
                    if let Some(tmp) = &result_temp {
                        stack.push(Expression::ident(&tmp));
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
                // ── rethrow N: reraise the exception this catch handler caught ─
                "rethrow" => {
                    let span = to_span(&pairs[i]);
                    let exnref = ACTIVE_CATCH_EXNREFS.with(|s| s.borrow().last().cloned());
                    match exnref {
                        Some(name) => statements.push(Statement::with_span(
                            StmtKind::WasmRethrow { exnref_local: name },
                            span,
                        )),
                        None => return Err("rethrow outside of a catch handler".to_string()),
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
                        let target = br_target_of(args.first());
                        if let Some(entry) = labels.resolve(&target) {
                            // Unconditional branch: carry the top of stack into a
                            // value-producing target, then jump.
                            if let Some(tmp) = &entry.result_temp {
                                if let Some(val) = stack.pop() {
                                    statements.push(Statement::new(StmtKind::Expr(
                                        Expression::new(ExprKind::Assign {
                                            target: Box::new(Expression::ident(tmp)),
                                            value: Box::new(val),
                                        }),
                                    )));
                                }
                            }
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
                                // The block result passes through a conditional
                                // branch, so peek (don't consume) the stack value.
                                if let Some(tmp) = &entry.result_temp {
                                    if let Some(val) = stack.last() {
                                        then_body.push(Statement::new(StmtKind::Expr(
                                            Expression::new(ExprKind::Assign {
                                                target: Box::new(Expression::ident(tmp)),
                                                value: Box::new(val.clone()),
                                            }),
                                        )));
                                    }
                                }
                                br_stmt_for(&entry, span)
                            }
                            None => make_br_stmt_opt(None, labels, span),
                        };
                        then_body.push(branch);
                        statements.push(Statement::with_span(
                            StmtKind::If {
                                cond: cond_expr,
                                then_body,
                                else_body: None,
                                elifs: Vec::new(),
                            },
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
                        let br_for = |t: &BrTarget| match labels.resolve(t) {
                            Some(entry) => br_stmt_for(&entry, span),
                            None => make_br_stmt_opt(None, labels, span),
                        };
                        if targets.is_empty() {
                            // Degenerate: nothing to branch to.
                        } else if targets.len() == 1 {
                            statements.push(br_for(&targets[0]));
                        } else {
                            let idx_tmp = fresh_result_temp();
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(idx_tmp.clone()),
                                    type_hint: None,
                                    init: Some(index),
                                    array_bounds: None,
                                    with_events: false,
                                }],
                                kind: VarDeclKind::Let,
                            }));
                            // Default (last) branch, then wrap each earlier case.
                            let mut chain = vec![br_for(&targets[targets.len() - 1])];
                            for k in (0..targets.len() - 1).rev() {
                                let cond = Expression::new(ExprKind::Binary {
                                    op: BinOp::StrictEq,
                                    left: Box::new(Expression::ident(&idx_tmp)),
                                    right: Box::new(Expression::int(k as i64)),
                                });
                                chain = vec![Statement::with_span(
                                    StmtKind::If {
                                        cond,
                                        then_body: vec![br_for(&targets[k])],
                                        else_body: Some(chain),
                                        elifs: Vec::new(),
                                    },
                                    span,
                                )];
                            }
                            statements.extend(chain);
                        }
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

        // Memory size / grow / bulk. fill/copy/init each pop 3 stack operands
        // (their data/mem-index selectors are immediates, not stack operands).
        "memory.size" => 0,
        "memory.grow" => 1,
        "memory.fill" | "memory.copy" | "memory.init" => 3,

        // Tables. The table index is an immediate; these are the stack operands.
        "table.get" => 1,             // elem index
        "table.set" | "table.grow" => 2, // (index,value) / (init,delta)
        "table.size" => 0,
        "table.fill" | "table.copy" | "table.init" => 3,

        // GC references without a type/field immediate — pure stack arity.
        "ref.i31" => 1,               // i32 → i31ref
        "i31.get_s" | "i31.get_u" => 1, // i31ref → i32
        "ref.as_non_null" | "any.convert_extern" | "extern.convert_any" => 1,
        "ref.is_null" => 1,           // [ref] → [i32]
        "ref.eq" => 2,

        // ── Stringref proposal (stack-operand counts; $mem is an immediate) ──
        "string.new_utf8" | "string.new_wtf8" | "string.new_lossy_utf8" => 2, // ptr, len
        "string.new_utf8_array" | "string.new_wtf16_array"
        | "string.new_wtf8_array" | "string.new_lossy_utf8_array" => 3, // arr, start, end
        "string.measure_utf8" | "string.measure_wtf8" | "string.measure_wtf16" => 1,
        "string.encode_utf8" | "string.encode_wtf16"
        | "string.encode_lossy_utf8" | "string.encode_wtf8" => 2, // str, ptr
        "string.encode_utf8_array" | "string.encode_wtf16_array"
        | "string.encode_lossy_utf8_array" | "string.encode_wtf8_array" => 3, // str, arr, start
        "string.concat" | "string.eq" | "string.compare" => 2,
        "string.is_usv_sequence"
        | "string.as_wtf8" | "string.as_wtf16" | "string.as_iter" => 1,
        "stringview_iter.next" | "stringview_iter.advance"
        | "stringview_wtf16.length" => 1,
        "array.len" => 1,             // arrayref → i32
        // Array ops carrying a type-index immediate (kept as an immediate arg):
        "array.new" => 2,            // value, length
        "array.new_default" => 1,    // length
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
        "array.get" => 2,            // arrayref, index
        "array.set" => 3,            // arrayref, index, value
        "array.fill" => 4,           // arrayref, index, value, count
        "array.copy" => 5,           // dst, dst_off, src, src_off, len (2 typeidxs dropped)

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

        // ── SIMD v128: number of STACK operands (lane index / v128.const values
        //    are immediates, not stack operands). ────────────────────────────
        n if is_simd_instr(n) => simd_stack_arity(n),

        _ => 0,
    }
}

/// Is this a SIMD (v128) instruction mnemonic?
fn is_simd_instr(name: &str) -> bool {
    matches!(name.split_once('.').map(|(p, _)| p), Some(
        "i8x16" | "i16x8" | "i32x4" | "i64x2" | "f32x4" | "f64x2" | "v128"
    ))
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

fn get_instruction_push_count(name: &str) -> usize {
    match name {
        "local.set" | "global.set" | "drop" | "br_if" | "br" | "unreachable" | "nop"
        | "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8" | "i32.store16"
        | "i64.store8" | "i64.store16" | "i64.store32" | "struct.set" => 0,
        _ => 1,
    }
}
