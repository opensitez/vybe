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
use std::collections::HashMap;
use vybe_ast::*;

/// Every registry the wast walk keeps, owned by one `parse` call.
///
/// These were 36 `thread_local!` statics, of which 34 were never cleared —
/// so in any host that compiles more than one program on a thread (the warm
/// worker pool, `--serve`, a runtime include) one script's module classes,
/// table/memory index bases and tag entities were still visible to the next.
/// A wast script legitimately accumulates across the MODULES of one script;
/// it must not accumulate across scripts, and a struct that `parse` owns is
/// what makes those two different things.
#[derive(Default)]
struct WastWalker {
    func_index_arities: Vec<usize>,
    func_name_arities: HashMap<String, usize>,
    func_index_results: Vec<usize>,
    func_name_results: HashMap<String, usize>,
    current_fn_results: usize,
    struct_field_counts: HashMap<String, usize>,
    struct_types: Vec<(String, Option<String>, usize)>,
    struct_field_types: HashMap<String, Vec<String>>,
    /// Struct type name → its field `$id`s, index-aligned with
    /// `struct_field_types` (empty string = an unnamed field). `struct.get $T $y`
    /// resolves through this; without it a named field became index 0.
    struct_field_ids: HashMap<String, Vec<String>>,
    /// Declared TYPE index → name, in declaration order. A type immediate may
    /// be written positionally (`struct.get_s 0 0`, `gc/struct.wast`), and the
    /// name is what every per-type map here is keyed by.
    type_index_name: Vec<String>,
    type_func_params: HashMap<String, usize>,
    /// Call Tags proposal: `(func … (call_tag $t+))` — func name → the tags
    /// its funcref handles. Collected while walking func fields and emitted as
    /// declarations once the module is complete, because the statement names
    /// the func and the func is what is being walked.
    func_call_tag_decls: Vec<(String, Vec<String>)>,
    /// Call Tags proposal: tag name → (params, results). The folder needs the
    /// tag's arity to know how many stack operands `call_with_tag` consumes,
    /// exactly as `call_ref` needs its `$sig`'s param count.
    call_tag_params: HashMap<String, (usize, usize)>,
    /// Local alias → external `module:name` for IMPORTED call tags, so a
    /// `call_with_tag $local` names the exporter's entity.
    call_tag_alias: HashMap<String, String>,
    /// Declared func type name → (param val types, result val types). A
    /// function type's identity is structural (`Comptype_sub/func`), so this
    /// is what `ref.test`/`ref.cast` against a concrete `(ref $t)` compares —
    /// arities alone would conflate `(param i32)` with `(param f64)`.
    type_func_sigs: HashMap<String, (Vec<String>, Vec<String>)>,
    /// Declared SUPERTYPE of a function type — `(type $sub (sub $super (func …)))`.
    /// Needed because two func types can share a signature and still be
    /// distinct: the subtype declaration is part of their identity.
    type_func_parent: HashMap<String, String>,
    type_func_results: HashMap<String, usize>,
    table_name_index: HashMap<String, usize>,
    memory_name_index: HashMap<String, usize>,
    /// THIS module's memory index space, mapped onto the script's: entry `i` is
    /// the script slot that module-relative memidx `i` names. An IMPORTED
    /// memory occupies an index here while pointing back at the exporter's
    /// slot, which is the whole reason a scalar base could not do the job.
    /// Declared type name → the name of the type it is CANONICALLY equal to.
    /// WASM 3.0 type identity is structural; see the note where this is built.
    type_canonical: HashMap<String, String>,
    /// Canonical type name → `"{group size}:{position}"` of its rec group.
    ///
    /// ⛔ THE ONE PART OF TYPE IDENTITY WE KNOW EXACTLY. Canonicalisation
    /// merges by the composite's source TEXT, which cannot see that
    /// `(param f32 f32)` and `(param $x f32) (param $y f32)` are one type, nor
    /// that `(ref $r1)` and `(ref $r2)` are when `$r1` and `$r2` are — so two
    /// equal types can end up with different names. The rec group's SIZE and
    /// the member's POSITION come from the tree instead, and under
    /// iso-recursive equivalence a difference in either means the types are
    /// genuinely different. That makes it the only safe discriminator for a
    /// runtime trap.
    type_rec_shape: HashMap<String, String>,
    memory_slots: Vec<usize>,
    /// One entry per `(memory …)` FIELD, in order: the script slot it owns and
    /// whether it merely aliases an imported memory. `walk_memory_field` counts
    /// fields, not index-space entries, so it cannot read `memory_slots`.
    memory_field_info: Vec<(usize, bool)>,
    /// Per module class: exported memory name → script slot, so a later
    /// module's `(memory (import "M" "mem"))` can find what it is aliasing.
    module_memory_exports: HashMap<String, HashMap<String, usize>>,
    table_index_base: usize,
    memory_index_base: usize,
    /// Data segments share ONE list across the whole script (the compiler
    /// pushes every `DataSegment`, active and passive alike, onto
    /// `chunks[0].data_segments`), but a written dataidx is MODULE-relative —
    /// exactly the split `memory_index_base` exists for. Without the shift, the
    /// second module's `memory.init 0` copied the FIRST module's segment, which
    /// is `bulk-memory/bulk.wast` and `memory64/bulk64.wast`.
    data_index_base: usize,
    data_name_index: HashMap<String, usize>,
    /// Element segments have the same script-wide-store / module-relative-index
    /// split as data segments: `__wast_register_passive_elem` keys one map for
    /// the whole program, so the second module's `table.init 0` looked up a
    /// segment the first module owned — reported as "missing element segment".
    elem_index_base: usize,
    elem_name_index: HashMap<String, usize>,
    module_seq: usize,
    current_module_seq: usize,
    /// Set by `walk_assert_trap`, so `parse` prepends the `__wast_check_trap`
    /// helper. Emitted only when something calls it — a wast script with no
    /// `assert_trap` at all gets no extra top-level function.
    needs_trap_contains: bool,
    func_index_name: Vec<String>,
    module_exports: HashMap<String, HashMap<String, String>>,
    export_global_map: HashMap<String, String>,
    module_global_exports: HashMap<String, HashMap<String, String>>,
    /// An imported global's local binding name → the EXPORTING module's
    /// binding for it. Globals share one top-level binding space, so aliasing
    /// the name is what makes an imported mutable global one shared cell
    /// rather than a copy.
    global_import_alias: HashMap<String, String>,
    global_index_name: Vec<String>,
    /// THIS function's local index space, in order: params first, then locals.
    /// A WAT local may be addressed by INDEX whatever it is called, so
    /// `local.get 0` has to find a `(param $a i32)` — the synthetic `p<i>`
    /// naming only ever worked for the UNNAMED case, and a named param made
    /// the numeric spelling read a binding that did not exist. Rebuilt per
    /// function, exactly as `current_fn_results` is.
    local_index_name: Vec<String>,
    array_elem_type: HashMap<String, String>,
    elem_seg_counter: usize,
    module_class_name: String,
    defined_func_names: std::collections::HashSet<String>,
    export_func_map: HashMap<String, String>,
    wast_temp_counter: usize,
    tag_arities: HashMap<String, u8>,
    tag_index_name: Vec<String>,
    tag_decl_ordinal: usize,
    tag_canon: HashMap<String, String>,
    module_tag_seq: usize,
    pending_tag_exports: Vec<(String, String)>,
    registered_tags: HashMap<(String, String), String>,
    registered_module_class: HashMap<String, String>,
    import_alias: HashMap<String, (String, String)>,
    /// A function import's local `$alias` → the `host:<module>:<fn>` callee
    /// spelled from the pair the import itself declares.
    ///
    /// WAT maps 1:1 onto WASM, and `(import "canon" "stream.read" (func $sr …))`
    /// already says everything needed to make the call. Before this map the
    /// walker kept the signature but threw the PAIR away, so the only way to
    /// reach a host function was for someone to pre-register the local alias in
    /// `languages/wast/src/profile` `[builtins]` — which is why `$log` resolved
    /// and `$stream_read` did not, why the alias had to be spelled exactly like
    /// the profile key, and why every importable function needed another row
    /// forever. That was the compiler discarding half the import, not WAT
    /// failing to be WASM.
    ///
    /// `Profile::lookup_builtin` synthesises the definition for a `host:`
    /// callee, so nothing has to be declared anywhere.
    host_import_alias: HashMap<String, String>,

    /// The component's TYPE space — `$id -> typeidx`, plus its length, so an
    /// unnamed `(type …)` still advances the index every later row counts on.
    /// Separate from `core_type_space` for the reason `CalleeRef` keeps `Core`
    /// and `Component` apart: they are different index spaces, and one integer
    /// serving both is how a mis-wired definition ends up looking correct.
    comp_type_index: HashMap<String, u32>,
    comp_type_space: u32,
    /// Where THIS component's type index 0 sits in the shared `comp_types`
    /// payload vector.
    ///
    /// ⛔ `comp_type_space` is a per-component COUNT (the spec gives every
    /// component its own type index space starting at 0) while `comp_types` is
    /// one flat vector for the whole program, because `VM::canon_types` is one
    /// table. Those two cannot both be right without an offset: a nested
    /// component that declared a type after the outer had declared one reset
    /// the counter to 0 while the vector kept growing, so the inner's
    /// `(type $t)` got index 0 and its declaration sat at index 1.
    ///
    /// In a debug build that tripped the `debug_assert_eq!` below. In a release
    /// build it was SILENT: `canon lift (type $t)` read the OUTER component's
    /// type and lifted with a signature the source never wrote.
    comp_type_base: u32,
    /// The component's FUNCTION space — what `canon lower` indexes. Not the
    /// type space: `(type $t)` and `(func $t)` are different entities.
    comp_func_index: HashMap<String, u32>,
    /// That space itself. `canon lift` is its producer: the spec says a lift
    /// DEFINES a component function, and `canon lower`'s callee indexes here.
    /// Until it existed `lower` could not name a callee and could not learn
    /// its type — `Binary.md:297` gives the lower row NO `ft` immediate,
    /// because `canon_lower(callee, ft, opts, …)` takes `ft` FROM THE CALLEE.
    comp_func_space: Vec<CompFunc>,
    /// The component's INSTANCE index space, in declaration order, plus its
    /// names. `(alias export <instanceidx> …)` resolves against this the way
    /// `(alias core export …)` resolves against `core_instances`; the two are
    /// deliberately separate spaces, because `(instance 0)` at component level
    /// and `(core instance 0)` name different entities.
    comp_instances: Vec<CompInstance>,
    comp_instance_index: HashMap<String, u32>,
    /// The component's own EXPORT table — `(export "name" (func $f))`.
    ///
    /// `Explainer.md:889`: a component type "contains *two* lists of named
    /// definitions for the imports and exports of a component". Instantiating
    /// a component yields an instance whose exports are exactly this list, so
    /// this is the same shape as `CompInstance::funcs` on purpose — it IS the
    /// export table of the instance `(instance (instantiate <componentidx>))`
    /// produces.
    ///
    /// ⛔ That instantiation does not exist yet, so nothing reads this today.
    /// It is recorded rather than dropped because the alternative is to walk
    /// an `(export …)` for its index-space effect and silently lose the name,
    /// and because the consumer is the very next item on the list — not an
    /// indefinite park. Say so if it is still unread by the time anything else
    /// lands.
    comp_exports: HashMap<String, u32>,
    /// The names each scope has taken, for the STRONGLY-UNIQUE rule. Two
    /// tables because `Explainer.md:2837` scopes the rule to imports OR
    /// exports, not both together.
    export_names: ExternNames,
    import_names: ExternNames,
    /// The component's CORE index spaces. Each is its OWN map, never one map
    /// consulted for several sorts: `(core type $t)` and `(core func $t)` may
    /// both be bound, to different entities, and a shared map would silently
    /// answer one for the other. They stay empty until `(core instance …)` and
    /// `(alias …)` are walked, and an unbound `$id` in one of them refuses.
    core_type_index: HashMap<String, u32>,
    core_type_space: u32,
    core_func_index: HashMap<String, u32>,
    core_memory_index: HashMap<String, u32>,
    core_table_index: HashMap<String, u32>,
    /// The core function index space ITSELF. `core_func_index` names entries
    /// in here; a positional `(core func 0)` indexes it directly. Both
    /// spellings had a name map and no space to name, which is why every one
    /// of them refused: the maps were read in six places and written in none.
    core_func_space: Vec<CoreFunc>,
    /// The core INSTANCE index space, in declaration order, plus its names.
    /// `(alias core export <instanceidx> …)` is the only way a component
    /// reaches inside an instantiated module, so without this space there is
    /// nothing for an alias to resolve against.
    core_instances: Vec<CoreInstance>,
    core_instance_index: HashMap<String, u32>,
    /// The component's CANON SECTION, in canonidx order.
    canon_section: Vec<vybe_ast::canon::CanonDecl>,
    /// The component's TYPE index space, in typeidx order. `walk_comp_type`
    /// used to advance a bare counter; the entries themselves had nowhere to
    /// go, which is why `VM::canon_functypes` was empty and `canon lift`
    /// trapped on `$ft` even when the source declared the type.
    comp_types: Vec<vybe_ast::canon::TypeDecl>,
    /// A canon row's `(core func $id)` binder → `(spec name, canonidx)`.
    ///
    /// This is what makes a canon definition REACHABLE. Without it a core
    /// module could only address a built-in by hand-writing the canonidx into
    /// the import name (`"thread.spawn-ref@0"`), which is the addressing
    /// deviation the canon section exists to remove.
    canon_binder: HashMap<String, (String, u32)>,
    /// `(import <module> <name>)` → the callee an INSTANTIATION supplies for
    /// it, from a `(with <module> (instance (export <name> (func $b))))`
    /// clause. Populated for the duration of one `(core instance …)` walk and
    /// cleared after, because it is that instantiation's wiring and not the
    /// module's — the same module instantiated twice may be given different
    /// imports, which is the entire point of `instantiate`.
    component_imports: HashMap<(String, String), CoreFunc>,
}

/// One entry in a component's CORE FUNCTION index space.
///
/// A component's core function is not always a module's function: a canonical
/// definition defines one too (`canon lower`, and every canonical built-in).
/// Both live in ONE space because `canon lift`'s `<core:funcidx>` and a `with`
/// clause's `(core func <idx>)` index it positionally and cannot tell which
/// kind they are naming — so the discriminant has to travel with the entry
/// rather than with the reference.
#[derive(Clone, Debug, PartialEq)]
enum CoreFunc {
    /// A canon row's core function: `(spec name, canonidx)`.
    Canon(String, u32),
    /// A function exported by an instantiated core module: `(class, method)`.
    Module(String, String),
}

/// One entry in a component's FUNCTION index space.
///
/// `canonidx` is the row that DEFINES this function — a `canon lift`, or the
/// entry an `(alias export …)` / `(export …)` copies. `functype` is that row's
/// `$ft`, carried here so a `canon lower` naming this function takes its type
/// rather than re-deriving it: the two must agree by construction, not by both
/// looking it up.
///
/// ⛔ `canonidx` is an `Option` because of `(import "x" (func $x (type $ft)))`.
/// An IMPORTED component function occupies an index in declaration order and
/// has no defining row anywhere in this component — so the slot must exist and
/// must be empty. Skipping the slot instead would renumber every function
/// declared after the import, which reads as correct until two of them share a
/// signature.
#[derive(Clone, Debug)]
struct CompFunc {
    canonidx: Option<u32>,
    functype: Option<u32>,
}

/// What a `(core instance …)` published: its export table.
///
/// The values are `CoreFunc` rather than method names because an instance
/// assembled from `<core:inlineexport>*` may export a CANON row, which has no
/// method — and an alias into it must get back the same kind of item a `with`
/// clause would.
#[derive(Clone, Debug, Default)]
struct CoreInstance {
    funcs: HashMap<String, CoreFunc>,
}

/// What an `(instance …)` published: its export table, at COMPONENT level.
///
/// A component instance exports component items, so the values are indices
/// into `comp_func_space` rather than `CoreFunc`s — a component instance
/// cannot export a core function, and conflating the two is exactly the
/// index-space confusion the separate spaces exist to prevent.
///
/// Only the function sort is present because only the function sort has a
/// producer today. A `(export "t" (type 0))` refuses rather than being
/// dropped, so the export table can never be quietly incomplete.
#[derive(Clone, Debug, Default)]
struct CompInstance {
    funcs: HashMap<String, u32>,
}


/// Payload arity of exception tag `name` (0 if undeclared).
fn tag_arity(__w: &mut WastWalker, name: &str) -> u8 {
    __w.tag_arities.get(name).copied().unwrap_or(0)
}

/// The canonical name a `throw`/`catch` tag reference denotes. `$e` is the name
/// itself; a bare number is a `tagidx` into the module's tag index space, which
/// must resolve to the SAME name the declaration registered — otherwise the
/// compiler imports two different entities for one tag and nothing matches.
fn tag_ref_name(__w: &mut WastWalker, raw: &str) -> String {
    let local = match raw.strip_prefix('$') {
        Some(name) => name.to_string(),
        None => match raw.parse::<usize>() {
            Ok(idx) => __w.tag_index_name.get(idx).cloned()
                .unwrap_or_else(|| raw.to_string()),
            Err(_) => raw.to_string(),
        },
    };
    __w.tag_canon.get(&local).cloned()
        .unwrap_or(local)
}

/// A fresh unique identifier for a structured-control result temporary.
fn fresh_result_temp(__w: &mut WastWalker) -> String {
    let name = format!("__wat_res{}", __w.wast_temp_counter);
    __w.wast_temp_counter += 1;
    name
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

/// How many result values a `block`/`loop`/`if`/`try` opener yields — the total
/// `any_val_type` count across its `(result …)` block-type immediates, plus any
/// `(type $sig)` immediate's result count. 0 = void, 1 = single-value baseline,
/// N = WASM multi-value.
fn peek_block_result_count(__w: &mut WastWalker, pair: &Pair<Rule>) -> usize {
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
                count += __w.type_func_results.get(&qualify_type_name(__w, &name)).copied()
                    .unwrap_or(0);
            }
        }
    }
    count
}

/// Does any NESTED operand of this folded instruction produce several values?
/// Only a `call` to a multi-result callee does, and only that case needs the
/// operand-spreading path.
fn has_multi_result_call_operand(__w: &mut WastWalker, op: &Pair<Rule>) -> bool {
    let children: Vec<Pair<Rule>> = op.clone().into_inner().collect();
    for child in children {
        let Some(nested) = folded_operand_child(&child) else {
            continue;
        };
        if nested.as_rule() != Rule::folded_instr {
            continue;
        }
        if folded_instr_head(&nested) != "call" {
            continue;
        }
        let imms: Vec<Expression> = nested
            .clone()
            .into_inner()
            .filter(|c| c.as_rule() == Rule::instr_arg && folded_operand_child(c).is_none())
            .filter_map(|c| walk_instr_arg_pair(__w, c, &mut LabelStack(Vec::new())).ok())
            .collect();
        if call_result_count(__w, &imms) >= 2 {
            return true;
        }
    }
    false
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
                value: Box::new(value),
            }));
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
                    value: Box::new(val.clone()),
                },
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
fn preserve_stack_across_block(__w: &mut WastWalker, stack: &mut Vec<Expression>, statements: &mut Vec<Statement>) {
    let pending: Vec<Expression> = stack.drain(..).collect();
    for e in pending {
        // A bare value expression (const/local.get/ident) has no side effect and
        // needs no temp — keep it deferred as-is. Anything else is bound to a
        // temp so its effect runs here, in order.
        let keep = matches!(e.kind, ExprKind::Lit(_) | ExprKind::Ident(_));
        if keep {
            stack.push(e);
        } else {
            let t = fresh_result_temp(__w);
            statements.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(t.clone()),
                    type_hint: None,
                    init: Some(e),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }));
            stack.push(Expression::ident(&t));
        }
    }
}

/// Lower a TOP-LEVEL folded `(block …)`/`(loop …)` as a STATEMENT (mirrors the
/// unfolded block handler), so its body actually runs — `walk_folded_core`'s
/// block path only returned the trailing expression and DISCARDED the body,
/// which silently dropped side effects and any `br`/`br_on_*` inside. Leaves the
/// block's N result values on `stack` for the continuation.
fn emit_folded_block(__w: &mut WastWalker, 
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
    let seed_vals = if param_count > 0 && stack.len() >= param_count {
        stack.split_off(stack.len() - param_count)
    } else {
        Vec::new()
    };
    preserve_stack_across_block(__w, stack, statements);
    // A `loop (param …)` threads its operand-stack params ACROSS ITERATIONS, so
    // each needs a synthetic local that the back-edge `br` can reassign — the
    // same model the unfolded handler uses. Seeding the body with the raw entry
    // EXPRESSIONS instead left `param_temps` empty, so `br 0` fell through to
    // the RESULT temps (see `LabelStack`'s loop branch) and the parameter never
    // advanced: `param-break` (loop.wast:301) stayed at its entry value 1
    // forever and the file HUNG instead of counting 1→5→9→13.
    let loop_param_temps: Vec<String> = if is_loop && param_count > 0 {
        (0..param_count).map(|_| fresh_result_temp(__w)).collect()
    } else {
        Vec::new()
    };
    let seed: Vec<Expression> = if loop_param_temps.is_empty() {
        seed_vals
    } else {
        for (k, tmp) in loop_param_temps.iter().enumerate() {
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
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }));
        }
        loop_param_temps
            .iter()
            .map(|t| Expression::ident(t))
            .collect()
    };
    let result_temps: Vec<String> = (0..result_count).map(|_| fresh_result_temp(__w)).collect();
    for tmp in &result_temps {
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
    let kind = if is_loop {
        LabelKind::Loop
    } else {
        LabelKind::Block
    };
    let effective = labels.push(__w, label.clone(), kind, result_temps.clone());
    // Publish the param temps so a back-edge `br` assigns the NEXT iteration's
    // values into them rather than into the result temps.
    labels.set_last_param_temps(loop_param_temps);
    let mut body = fold_instructions_seeded(__w, instr_pairs, labels, seed)?;
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
    for tmp in &result_temps {
        stack.push(Expression::ident(tmp));
    }
    Ok(())
}

/// Statement lowering for a folded `(if $label? bt? (cond)* (then …) (else …)?)`.
/// The expression walk lowers folded `if` to a ternary that keeps only each
/// arm's LAST expression — a `br`/`br_if`/`return` inside an arm is a
/// STATEMENT and gets dropped there (labels.wast `loop1`: the `(br $exit …)`
/// inside `(then …)` vanished, so the loop never exited). This mirrors the
/// unfolded-if handler: real If statement, arms folded recursively with the
/// full stack machinery, results captured in temps.
/// Is this `instr_arg`/`instr` child a nested OPERAND (an instruction that
/// produces a value) rather than an immediate?
fn folded_operand_child<'a>(child: &Pair<'a, Rule>) -> Option<Pair<'a, Rule>> {
    match child.as_rule() {
        Rule::instr => child.clone().into_inner().next(),
        Rule::instr_arg => child
            .clone()
            .into_inner()
            .next()
            .filter(|x| matches!(x.as_rule(), Rule::folded_instr | Rule::plain_instr)),
        _ => None,
    }
}

/// Evaluate ONE folded operand and leave its value on the stack.
///
/// Folding is sugar: `(instr o1 … ok)` unfolds to `o1 … ok instr`, and the
/// spec's abbreviation lets a folded instruction nest FEWER operands than it
/// consumes — the rest come from the enclosing stack, BELOW the nested ones.
/// `(i32.ne)` after two pushes and `(i32.ne (i32.const 2))` after one are both
/// legal, and both appear in the spec suite (`throw.wast`'s `test-throw-1-2`
/// compares the two values a `catch` delivered).
///
/// The expression walk cannot express that: `walk_folded_instr_as_expr` sees
/// only the nested operands and substitutes defaults for the rest, so
/// `(if (i32.ne) …)` compared 0 with 0 and the branch was constant. Only the
/// fully-nested case is left to it.
fn push_folded_operand(__w: &mut WastWalker, 
    op: Pair<Rule>,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    if folded_needs_stmt_lowering(&op) {
        return emit_folded_operand_stmtwise(__w, op, labels, statements, stack);
    }
    let op_span = to_span(&op);
    let head = folded_instr_head(&op);
    let nested_count = op
        .clone()
        .into_inner()
        .filter(|c| folded_operand_child(c).is_some())
        .count();
    let immediate_args: Vec<Expression> = op
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::instr_arg && folded_operand_child(c).is_none())
        .map(|c| walk_instr_arg_pair(__w, c, labels))
        .collect::<Result<_, _>>()?;
    let arity = get_instruction_arity(__w, &head, &immediate_args);
    // ⛔ A NESTED MULTI-RESULT CALL MUST TAKE THIS PATH TOO. It is the only one
    // that evaluates operands through `push_folded_operand`, which spreads a
    // call's several results into several stack entries; the expression-only
    // fallback below collapses them into one argument. `(call $swap (call
    // $swap …))` has nested_count 1 < arity 2 but an EMPTY stack, so the
    // `!stack.is_empty()` guard sent it to the fallback and the outer call
    // received one packed tuple instead of two operands.
    //
    // Entering with nothing below is safe: `take` is then 0 and the drain is a
    // no-op — the guard was avoiding a pointless drain, not preventing one.
    let spread_operand = has_multi_result_call_operand(__w, &op);
    if nested_count < arity && (!stack.is_empty() || spread_operand) {
        // The operands the flat form would have found already on the stack sit
        // BELOW the nested ones — take them first, then evaluate the nested
        // operands in source order on top.
        let take = (arity - nested_count).min(stack.len());
        let drain_start = stack.len() - take;
        let below: Vec<Expression> = stack.drain(drain_start..).collect();
        let base = stack.len();
        for child in op.clone().into_inner() {
            if let Some(nested) = folded_operand_child(&child) {
                if nested.as_rule() == Rule::folded_instr {
                    push_folded_operand(__w, nested, labels, statements, stack)?;
                } else {
                    let s = to_span(&nested);
                    stack.push(walk_plain_instr_as_expr(__w, nested, s, labels)?);
                }
            }
        }
        let nested_vals: Vec<Expression> = stack.drain(base..).collect();
        let mut args = immediate_args;
        args.extend(below);
        args.extend(nested_vals);
        // ⛔ A MULTI-RESULT CALL PUSHES SEVERAL OPERANDS. `land_instr_value`
        // already destructures them into one stack entry each — this path just
        // never reached it, so `(i32.add (call $swap-i32-i32 …))` handed
        // `i32.add` a single packed tuple where the spec gives it two values.
        let pushes = if head == "call" {
            call_result_count(__w, &args)
        } else {
            1
        };
        let value = map_instr_to_ast(__w, head.clone(), args, op_span)?;
        if pushes >= 2 {
            land_instr_value(__w, value, pushes, true, op_span, statements, stack);
        } else {
            stack.push(value);
        }
        return Ok(());
    }
    // ⛔ THE FULLY-NESTED PATH NEEDS THE SPREAD TOO — and it is the one a
    // multi-result call actually takes. `(call $swap-i32-i32 (i32.const 3)
    // (i32.const 4))` has nested_count == arity, so it never enters the branch
    // above and landed here, pushing ONE packed tuple where the callee's two
    // results are two operands.
    let pushes = if head == "call" {
        call_result_count(__w, &immediate_args)
    } else {
        1
    };
    let value = walk_folded_instr_as_expr(__w, op, op_span, labels)?;
    if pushes >= 2 {
        land_instr_value(__w, value, pushes, true, op_span, statements, stack);
    } else {
        stack.push(value);
    }
    Ok(())
}

fn emit_folded_if(__w: &mut WastWalker, 
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let mut label: Option<String> = None;
    let mut result_count = 0usize;
    let mut param_count = 0usize;
    let mut then_pairs: Vec<Pair<Rule>> = Vec::new();
    let mut else_pairs: Option<Vec<Pair<Rule>>> = None;
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
            Rule::then_block => {
                then_pairs = child
                    .into_inner()
                    .filter(|s| s.as_rule() == Rule::instr)
                    .collect();
            }
            Rule::else_block => {
                else_pairs = Some(
                    child
                        .into_inner()
                        .filter(|s| s.as_rule() == Rule::instr)
                        .collect(),
                );
            }
            // Folded condition operand(s) — evaluate onto the stack in order,
            // routing branchy/structured ones through the statement machinery.
            Rule::instr_arg => {
                let nested = child
                    .clone()
                    .into_inner()
                    .next()
                    .filter(|x| x.as_rule() == Rule::folded_instr);
                if let Some(op) = nested {
                    push_folded_operand(__w, op, labels, statements, stack)?;
                }
            }
            Rule::instr => {
                let nested = child.clone().into_inner().next();
                match nested {
                    Some(op) if op.as_rule() == Rule::folded_instr => {
                        push_folded_operand(__w, op, labels, statements, stack)?;
                    }
                    _ => stack.push(walk_instr_as_expr(__w, child, labels)?),
                }
            }
            _ => {}
        }
    }
    let cond = stack.pop().unwrap_or(Expression::bool(false));
    // Seed `(param …)` inputs into BOTH arms, then flush pending side effects
    // (mirrors the unfolded-if handler).
    let seed = if param_count > 0 && stack.len() >= param_count {
        stack.split_off(stack.len() - param_count)
    } else {
        Vec::new()
    };
    preserve_stack_across_block(__w, stack, statements);
    let result_temps: Vec<String> = (0..result_count).map(|_| fresh_result_temp(__w)).collect();
    for tmp in &result_temps {
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
    let effective = labels.push(__w, label.clone(), LabelKind::Block, result_temps.clone());
    let mut then_body = fold_instructions_seeded(__w, then_pairs, labels, seed.clone())?;
    assign_last_n_exprs_to(&mut then_body, &result_temps);
    let else_body = match else_pairs {
        Some(pairs) => {
            let mut eb = fold_instructions_seeded(__w, pairs, labels, seed.clone())?;
            assign_last_n_exprs_to(&mut eb, &result_temps);
            Some(eb)
        }
        // An `if` with NO `(else)` is valid only when the block type's params
        // and results MATCH, and the absent else is the IDENTITY — it passes
        // the params straight through as the results. Emitting `None` left the
        // result temps at their `null` init, so `params-id` (if.wast:426,
        // `(if (param i32 i32) (result i32 i32) … (then))`) yielded null on the
        // false branch instead of its two params (if.wast:667, cond 0).
        None if !result_temps.is_empty() => Some(
            seed.iter()
                .zip(result_temps.iter())
                .map(|(v, tmp)| {
                    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::ident(tmp)),
                        value: Box::new(v.clone()),
                    })))
                })
                .collect(),
        ),
        None => None,
    };
    labels.pop();
    let if_stmt = Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body,
            elifs: Vec::new(),
        },
        span,
    );
    statements.push(Statement::with_span(
        StmtKind::Labeled {
            label: effective,
            body: Box::new(if_stmt),
        },
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
fn emit_folded_try_table(__w: &mut WastWalker, 
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let mut clauses: Vec<Pair<Rule>> = Vec::new();
    let mut instr_pairs: Vec<Pair<Rule>> = Vec::new();
    let mut result_count = 0usize;
    let mut param_count = 0usize;
    let mut own_label: Option<String> = None;
    for child in inner.into_inner() {
        match child.as_rule() {
            Rule::try_clause => clauses.push(child),
            Rule::instr => instr_pairs.push(child),
            Rule::block_type => {
                let text = child.as_str().trim_start();
                let n = child
                    .into_inner()
                    .filter(|p| matches!(p.as_rule(), Rule::any_val_type | Rule::val_type))
                    .count();
                // `try_table bt …` carries a full blocktype: params as well as
                // results. Only results were counted, so `(try_table (param i32)
                // …)` silently lost its operand.
                if text.starts_with("(result") {
                    result_count += n;
                } else if text.starts_with("(param") {
                    param_count += n;
                }
            }
            // The try_table's OWN label. Clauses do not use it — they are
            // resolved in the enclosing context — but the BODY does: spec
            // "try_table acts as a regular block for br, etc.", so `br 0`
            // inside the body exits the try_table.
            Rule::id => own_label = Some(child.as_str()[1..].to_string()),
            _ => {}
        }
    }

    // Side effects pending on the stack must run before the protected region.
    preserve_stack_across_block(__w, stack, statements);

    // On NORMAL completion (nothing thrown) the body's trailing values are the
    // try_table's results, captured in temps left on the stack afterwards.
    let result_temps: Vec<String> = (0..result_count).map(|_| fresh_result_temp(__w)).collect();
    for tmp in &result_temps {
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
    // `try_table` IS a block, so its body is folded with its OWN label in scope
    // — `br 0` there exits the try_table, carrying values into the same result
    // temps normal completion uses. The label is popped before the clauses are
    // resolved: a clause target is named in the ENCLOSING context, where depth
    // 0 is the block around the try_table, not the try_table itself.
    let own_effective = labels.push(__w, own_label, LabelKind::Block, result_temps.clone());
    let mut body = fold_instructions(__w, instr_pairs, labels)?;
    labels.pop();
    assign_last_n_exprs_to(&mut body, &result_temps);

    let mut wasm_catches: Vec<WasmCatch> = Vec::new();
    for clause in clauses {
        let kw = clause.as_str();
        // Keep the `$` so a NUMERIC clause target stays distinguishable from a
        // named one: `(catch $e0 0)` names a relative block DEPTH, resolved the
        // way `br 0` is. Stripping it first made every numeric target look like
        // a label named "0", which resolved against nothing.
        let idxs: Vec<String> = clause
            .into_inner()
            .filter(|c| c.as_rule() == Rule::index)
            .map(|c| c.as_str().to_string())
            .collect();
        // catch/catch_ref: [tag, label]; catch_all/catch_all_ref: [label].
        let (tag, capture_ref, label) = if kw.starts_with("(catch_all_ref") {
            (None, true, idxs.first().cloned().unwrap_or_default())
        } else if kw.starts_with("(catch_all") {
            (None, false, idxs.first().cloned().unwrap_or_default())
        } else if kw.starts_with("(catch_ref") {
            (
                idxs.first().map(|t| tag_ref_name(__w, t)),
                true,
                idxs.get(1).cloned().unwrap_or_default(),
            )
        } else {
            (
                idxs.first().map(|t| tag_ref_name(__w, t)),
                false,
                idxs.get(1).cloned().unwrap_or_default(),
            )
        };

        let arity = tag.as_deref().map(|t| tag_arity(__w, t)).unwrap_or(0);
        let payload_binds: Vec<String> = (0..arity).map(|_| fresh_result_temp(__w)).collect();
        let exnref_bind = if capture_ref {
            Some(fresh_result_temp(__w))
        } else {
            None
        };

        // Handler ≡ `br $L` carrying the delivered payload (+exnref): assign them
        // into the target's carry temps, then branch. The compiler binds the
        // caught payload/exnref into these locals before running this body.
        let target = match label.strip_prefix('$') {
            Some(name) => BrTarget::Named(name.to_string()),
            None => match label.parse::<usize>() {
                Ok(depth) => BrTarget::Index(depth),
                Err(_) => BrTarget::Named(label.clone()),
            },
        };
        let mut hstack: Vec<Expression> =
            payload_binds.iter().map(|n| Expression::ident(n)).collect();
        if let Some(e) = &exnref_bind {
            hstack.push(Expression::ident(e));
        }
        let mut hbody: Vec<Statement> = Vec::new();
        match labels.resolve(&target) {
            Some(entry) => {
                let carry = branch_carry_temps(&entry);
                carry_stack_into_temps(&carry, &mut hstack, true, &mut hbody);
                hbody.push(br_stmt_for(&entry, span));
            }
            // Depth `labels.len()` is the FUNCTION's own implicit label — the
            // one `br` uses to return. `(func (try_table (catch $e 0) …))` with
            // no enclosing block names exactly that, and the delivered values
            // become the function's results.
            None if matches!(target, BrTarget::Index(d) if d == labels.len()) => {
                if hstack.len() >= 2 {
                    let n = hstack.len();
                    hbody.push(multi_value_return_stmt(&mut hstack, n, span));
                } else {
                    hbody.push(Statement::with_span(StmtKind::Return(hstack.pop()), span));
                }
            }
            None => {
                return Err(format!("try_table clause targets unknown label {label}"));
            }
        }

        wasm_catches.push(WasmCatch {
            tag,
            payload_binds,
            capture_ref,
            exnref_bind,
            body: hbody,
        });
    }

    // Wrapped in its own label so a `br` from the body has somewhere to land:
    // breaking the labelled statement exits the try_table and continues after
    // it, which is what a branch to a block's label means.
    let try_stmt = Statement::with_span(
        StmtKind::WasmTryTable {
            body,
            catches: wasm_catches,
            params: param_count as u8,
            results: result_count as u8,
        },
        span,
    );
    statements.push(Statement::with_span(
        StmtKind::Labeled {
            label: own_effective,
            body: Box::new(Statement::with_span(StmtKind::Block(vec![try_stmt]), span)),
        },
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
fn emit_folded_br_on_null(__w: &mut WastWalker, 
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
                    label_arg = Some(walk_instr_arg_pair(__w, child, labels)?);
                } else {
                    let e = walk_instr_arg_pair(__w, child, labels)?;
                    stack.push(e);
                }
            }
            Rule::instr => {
                let e = walk_instr_as_expr(__w, child, labels)?;
                stack.push(e);
            }
            _ => {}
        }
    }
    let ref_val = stack.pop().unwrap_or_else(Expression::null);
    let tmp = fresh_result_temp(__w);
    statements.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(tmp.clone()),
            type_hint: None,
            init: Some(ref_val),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    let is_null = make_call("ref_is_null", vec![Expression::ident(&tmp)], span);
    let target = br_target_of(label_arg.as_ref());
    let mut then_body: Vec<Statement> = Vec::new();
    match labels.resolve_dest(&target) {
        // ⛔ THE SAME FUNCTION-LABEL CASE AS `br`. `br_on_null 0` in a
        // function with no enclosing block names the function's own label and
        // must RETURN — it fell through here exactly as `br` did, and in the
        // same file, after the fix whose whole point was that sibling
        // spellings agree. Nothing in the suite showed it: the GC
        // `br_on_cast` fixtures stop on an earlier validation gap and
        // `br_on_non_null.wast` fails on a different bug of its own.
        BrDest::Func => {
            let n = __w.current_fn_results;
            // `br_on_non_null` delivers `t*` AND the ref (the ref is the top
            // result); `br_on_null` delivers only the `t*` below it. Peek in
            // both cases — the fall-through path still owns the stack.
            if is_non_null {
                let mut vals: Vec<Expression> = Vec::new();
                let below = n.saturating_sub(1).min(stack.len());
                vals.extend(stack[stack.len() - below..].iter().cloned());
                vals.push(Expression::ident(&tmp));
                then_body.push(func_return_stmt_n(n, &mut vals, true, span));
            } else {
                let mut vals = stack.clone();
                then_body.push(func_return_stmt_n(n, &mut vals, true, span));
            }
        }
        BrDest::Frame(entry) => {
            if is_non_null {
                // Carry the ref into the target's topmost result, then branch.
                if let Some(rt) = entry.result_temps.last() {
                    then_body.push(Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Assign {
                            target: Box::new(Expression::ident(rt)),
                            value: Box::new(Expression::ident(&tmp)),
                        },
                    ))));
                }
            } else {
                // br_on_null carries the values BELOW the ref (peeked).
                carry_stack_into_temps(&entry.result_temps, stack, false, &mut then_body);
            }
            then_body.push(br_stmt_for(&entry, span));
        }
        BrDest::Unresolved => then_body.push(make_br_stmt_opt(None, labels, span)),
    }
    let cond = if is_non_null {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(is_null),
            right: Box::new(Expression::int(0)),
        })
    } else {
        is_null
    };
    statements.push(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body: None,
            elifs: Vec::new(),
        },
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
fn emit_folded_return(__w: &mut WastWalker, 
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    for child in inner.into_inner() {
        match child.as_rule() {
            // A folded return value nests as `instr_arg → folded_instr`. A
            // branching/structured value (`(return (br_if …))`, `(return
            // (block …))`) must go through the statement machinery — the
            // expression walk would drop the branch/label handling.
            Rule::instr_arg => {
                let nested = child
                    .clone()
                    .into_inner()
                    .next()
                    .filter(|x| x.as_rule() == Rule::folded_instr);
                match nested {
                    // `push_folded_operand` routes stmt-lowering cases onward
                    // AND applies the spec's partial-fold rule: an operand may
                    // nest fewer values than it consumes, taking the rest from
                    // the enclosing stack. Evaluating it directly skipped that,
                    // so `(i32.const 7) (return (i32.add (i32.const 100)))`
                    // answered 100 while the unfolded form gave 107.
                    Some(op) => push_folded_operand(__w, op, labels, statements, stack)?,
                    None => stack.push(walk_instr_arg_pair(__w, child, labels)?),
                }
            }
            Rule::instr => {
                let nested = child.clone().into_inner().next();
                match nested {
                    Some(op) if op.as_rule() == Rule::folded_instr => {
                        push_folded_operand(__w, op, labels, statements, stack)?;
                    }
                    _ => stack.push(walk_instr_as_expr(__w, child, labels)?),
                }
            }
            _ => {}
        }
    }
    let n = __w.current_fn_results;
    if n >= 2 {
        statements.push(multi_value_return_stmt(stack, n, span));
    } else {
        let val = stack.pop();
        statements.push(Statement::with_span(StmtKind::Return(val), span));
    }
    Ok(())
}

/// Does this folded instr need STATEMENT lowering rather than the expression
/// walk? True when it branches (`br`/`br_if`/`br_table`/`return` — a branch
/// has no expression value; the expression walk yields null and DROPS it),
/// when it is structured control (`block`/`loop`/`if`/`try_table` — the
/// expression walk keeps only the body's last expression, dropping inner
/// branches, labels, and result-temp carries), or when a nested operand
/// transitively needs it.
fn folded_needs_stmt_lowering(pair: &Pair<Rule>) -> bool {
    let head = folded_instr_head(pair);
    if matches!(
        head.as_str(),
        "br" | "br_if"
            | "br_table"
            | "return"
            | "block"
            | "loop"
            | "if"
            | "try_table"
            // The `(type $sig)` immediates are read off the PAIR (the
            // expression walk drops them and the call gets argc/results 0→0).
            | "call_indirect"
            | "return_call_indirect"
            // A tail call must reuse the frame; the expression walk emits an
            // ordinary call to an unqualified callee instead.
            | "return_call"
            | "return_call_ref"
            // `select` evaluates ALL THREE operands (it is not a conditional);
            // the expression walk's ternary would short-circuit their effects.
            | "select"
            // These BRANCH — the expression walk emits them as a bare opcode,
            // and `Op::BR_ON_NULL`'s `I16` ip-offset has no arm in the emitter's
            // fixed-immediate match, so it lands 2 bytes short and the VM
            // decodes the next instruction mid-stream ("Invalid opcode: 0x0021
            // …"). They are already claimed at the HEAD of a flat sequence; this
            // also claims them as a nested OPERAND, e.g. br_on_null.wast:6
            // `(return (call_ref $t (br_on_null $l (local.get $r))))`.
            | "br_on_null"
            | "br_on_non_null"
            // `br_on_cast`/`br_on_cast_fail` BRANCH, and the value they carry
            // becomes the target block's result. Only the flat (plain)
            // sequence claimed them, so a FOLDED one was walked as an
            // expression: the branch became an ordinary call, and the block's
            // trailing "result = last expression" assignment hoisted it to the
            // END of the body — after the `return` it was supposed to jump
            // over. `gc/br_on_cast.wast` is written entirely in folded form.
            | "br_on_cast"
            | "br_on_cast_fail"
            // Same story for the Custom Descriptors branching casts: they
            // BRANCH, so they must be claimed statement-wise in both the flat
            // and the folded spelling. The proposal's own suite writes them
            // folded throughout.
            | "br_on_cast_desc_eq"
            | "br_on_cast_desc_eq_fail"
            // `throw`/`throw_ref` DIVERGE and carry a tag immediate. The
            // expression walk has neither: it resolved the head off the VM
            // opcode table and emitted `Op::THROW` with a ZERO tag operand, so
            // every folded `(throw $t …)` raised tag index 0 whatever tag it
            // named — right only by accident, when some `catch $t` had already
            // imported that tag as the chunk's first. With a `catch_all`-only
            // `try_table` nothing imports a tag at all and index 0 does not
            // resolve ("unknown exception tag index 0").
            | "throw"
            | "throw_ref"
    ) {
        return true;
    }
    pair.clone().into_inner().any(|c| {
        let nested = match c.as_rule() {
            Rule::folded_instr => Some(c),
            Rule::instr_arg | Rule::instr => c
                .into_inner()
                .next()
                .filter(|x| x.as_rule() == Rule::folded_instr),
            _ => None,
        };
        nested.is_some_and(|n| folded_needs_stmt_lowering(&n))
    })
}

/// Land an instruction's value(s) the way the plain path does: a void
/// instruction is a statement in program order; one value goes on the stack
/// for its consumer; a multi-result call destructures into fresh temps (the
/// shared compiler's multi-value ABI) and leaves the temps on the stack.
fn land_instr_value(__w: &mut WastWalker, 
    expr: Expression,
    pushes: usize,
    is_call: bool,
    span: Span,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    if is_call && pushes >= 2 {
        let temps: Vec<String> = (0..pushes).map(|_| fresh_result_temp(__w)).collect();
        let pats: Vec<ArrayPatternElem> = temps
            .iter()
            .map(|t| ArrayPatternElem::Pattern(BindingPattern::Ident(t.clone()), None))
            .collect();
        statements.push(Statement::with_span(
            StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Destructure(
                    DestructurePattern::Array(pats),
                ))],
                value: expr,
                by_ref: false,
            },
            span,
        ));
        for t in &temps {
            stack.push(Expression::ident(t));
        }
    } else if pushes > 0 {
        stack.push(expr);
    } else {
        // A VOID instruction runs at its program position, so anything still
        // DEFERRED on the value stack must be evaluated BEFORE it — otherwise
        // the deferred expressions materialise at their consumer, after this
        // statement, and their side effects run out of order. stack.wast:138
        // `not-quite-a-tree` is exactly that: two `call $add_one_to_global`
        // results sit deferred, a VOID `call $add_one_to_global_and_drop` lands
        // ahead of them, and the adds then evaluate 2+3 = 5 instead of 1+2 = 3.
        // Bare literals/idents stay deferred (no effect to order).
        preserve_stack_across_block(__w, stack, statements);
        statements.push(Statement::with_span(StmtKind::Expr(expr), span));
    }
}

/// Route one nested folded OPERAND through the statement machinery: structured
/// control goes to its dedicated emitter (which pushes its result temps onto
/// the stack), branching instrs recurse through `emit_folded_stmtwise`.
fn emit_folded_operand_stmtwise(__w: &mut WastWalker, 
    op: Pair<Rule>,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let op_span = to_span(&op);
    match folded_instr_head(&op).as_str() {
        "block" => emit_folded_block(__w, op, false, op_span, labels, statements, stack),
        "loop" => emit_folded_block(__w, op, true, op_span, labels, statements, stack),
        "if" => emit_folded_if(__w, op, op_span, labels, statements, stack),
        "try_table" => emit_folded_try_table(__w, op, op_span, labels, statements, stack),
        // Same structured lowering the flat head position uses — it leaves the
        // fall-through ref on the stack, which is exactly what the consuming
        // instruction (`call_ref`, `drop`, …) then pops.
        h @ ("br_on_null" | "br_on_non_null") => emit_folded_br_on_null(__w, 
            op,
            h == "br_on_non_null",
            op_span,
            labels,
            statements,
            stack,
        ),
        _ => emit_folded_stmtwise(__w, op, op_span, labels, statements, stack),
    }
}

/// Statement-aware lowering for a folded instr that branches (see
/// `folded_contains_branch`). Folding is pure sugar for the flat instruction
/// sequence (spec: text format, abbreviations), so this walks the nested
/// operands in order through the SAME stack+statements machinery as the plain
/// form — a nested `(br_if …)` peeks the block-result carry and emits its
/// conditional branch; the head then takes its operands from the stack.
/// `br_on_cast $l <from> <to>` / `br_on_cast_fail`: branch to `$l` carrying the
/// reference as that block's result when the cast succeeds (fails), and leave it
/// on the stack for the continuation otherwise.
///
/// `args` are the instruction's IMMEDIATES in source order — label, source heap
/// type, target heap type — with the reference operand on `stack`. Both the flat
/// and the folded parse sites hand it exactly that, which is the point of it
/// being one function: the folded site had no lowering at all, so the branch
/// became an ordinary call and the enclosing block's "result = last expression"
/// assignment moved it to the END of the body.
fn emit_br_on_cast_stmt(__w: &mut WastWalker,
    args: &[Expression],
    is_fail: bool,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    let target = br_target_of(args.first());
    // The target reftype is module-qualified like every other type reference.
    // In the FOLDED spelling it arrives as a `ref(…)` call that
    // `walk_instr_arg_pair` already qualified; in the PLAIN one it is a bare
    // `$id` and needs it here.
    let to_ht = match args.get(2).map(|e| &e.kind) {
        Some(ExprKind::Ident(n)) => Expression::ident(&qualify_type_name(__w, n)),
        _ => args.get(2).cloned().unwrap_or(Expression::null()),
    };
    // Consume the ref into a temp: on branch it becomes the block result; on
    // fall-through it is pushed back so the continuation (e.g. `drop`) consumes
    // it. Binding once avoids re-evaluating and keeps the stack balanced on
    // both paths.
    let ref_val = stack.pop().unwrap_or_else(Expression::null);
    let tmp = fresh_result_temp(__w);
    statements.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(tmp.clone()),
            type_hint: None,
            init: Some(ref_val),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    let test = make_call("ref_test", vec![to_ht, Expression::ident(&tmp)], span);
    let cond = if is_fail {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(test),
            right: Box::new(Expression::int(0)),
        })
    } else {
        test
    };
    let mut then_body: Vec<Statement> = Vec::new();
    match labels.resolve_dest(&target) {
        BrDest::Frame(entry) => {
            // ⚠ THE BRANCH CARRIES EVERY RESULT, NOT JUST THE REFERENCE.
            //
            // The typing rule is `t* rt_1 -> t* (rt_1 \ rt_2)`, and that
            // leading `t*` is real: a `(block (result i32 i64 anyref) …)`
            // branched out of by `br_on_cast` must deliver the two numbers
            // sitting BELOW the reference as well. Assigning only the topmost
            // result left them unwritten, so the block delivered garbage —
            // a WRONG VALUE, never a validation error, which is why nothing
            // upstream caught it.
            //
            // The reference is the top result; the values below it come off
            // the stack, untouched and NOT consumed, because the fall-through
            // path still needs them. Identical to `br_on_cast_desc_eq`, which
            // got this treatment when the descriptor suite's "Sent values"
            // section exercised it. ⛔The plain form could not be caught the
            // same way: every block in the GC `br_on_cast` fixtures is
            // single-result, so `t*` is empty throughout and the drop is
            // invisible there.
            let temps = &entry.result_temps;
            if let Some((ref_temp, below)) = temps.split_last() {
                carry_stack_into_temps(below, stack, false, &mut then_body);
                then_body.push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Assign {
                        target: Box::new(Expression::ident(ref_temp)),
                        value: Box::new(Expression::ident(&tmp)),
                    },
                ))));
            }
            then_body.push(br_stmt_for(&entry, span));
        }
        // ⚠ THE FUNCTION-LEVEL LABEL IS A RETURN, NOT A BARE BREAK.
        //
        // `br_on_cast 0` written directly in a function body targets the
        // function's implicit block, and branching to it RETURNS carrying the
        // reference. Emitting `Break(Implicit)` left the value behind and fell
        // through to whatever followed — in the spec fixtures, an
        // `(unreachable)` placed there precisely because the branch should have
        // jumped over it. The same instruction inside an explicit `(block …)`
        // resolved to a real label and worked, so the two spellings disagreed.
        //
        // ⛔ AND IT CARRIES `t*` TOO, not just the reference. Returning the
        // ref alone is right only for a single-result function; the typing
        // rule's leading `t*` is as real here as it is on the block path
        // directly above, where assigning only the topmost result delivered
        // garbage for the values underneath.
        BrDest::Func => {
            let n = __w.current_fn_results;
            let below = n.saturating_sub(1).min(stack.len());
            let mut vals: Vec<Expression> = stack[stack.len() - below..].to_vec();
            vals.push(Expression::ident(&tmp));
            then_body.push(func_return_stmt_n(n, &mut vals, true, span));
        }
        BrDest::Unresolved => then_body.push(Statement::with_span(
            StmtKind::Return(Some(Expression::ident(&tmp))),
            span,
        )),
    }
    statements.push(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body: None,
            elifs: Vec::new(),
        },
        span,
    ));
    // Fall-through: the ref stays available for continuation.
    stack.push(Expression::ident(&tmp));
}

/// Is this `br_on_cast*` target reftype immediate NULLABLE?
///
/// `walk_instr_arg_pair` hands `(ref null $a)` over as a call to `ref` whose
/// first argument is the marker ident `null`, `(ref $a)` as the same call
/// without it, and a bare `anyref`/`structref` spelling as a string. The bare
/// spellings are the nullable shorthand, so only the explicit `(ref ht)` form
/// is non-nullable.
fn reftype_is_nullable(arg: Option<&Expression>) -> bool {
    match arg.map(|a| &a.kind) {
        Some(ExprKind::Call { callee, args, .. })
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "ref") =>
        {
            args.iter()
                .any(|a| matches!(&a.value.kind, ExprKind::Ident(n) if n == "null"))
        }
        // `anyref`, `structref`, … — the abstract shorthand, always nullable.
        _ => true,
    }
}

/// The heap type NAME out of a `br_on_cast*` reftype immediate: `$a` out of
/// `(ref null $a)`, and the bare spelling out of `anyref`.
fn reftype_heap_name(arg: Option<&Expression>) -> String {
    match arg.map(|a| &a.kind) {
        Some(ExprKind::Call { callee, args, .. })
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "ref") =>
        {
            args.iter()
                .rev()
                .find_map(|a| match &a.value.kind {
                    ExprKind::Ident(n) if n != "null" => Some(n.clone()),
                    ExprKind::Lit(Literal::Str(s)) => Some(s.to_string()),
                    _ => None,
                })
                .unwrap_or_default()
        }
        Some(ExprKind::Lit(Literal::Str(s))) => s.to_string(),
        // The PLAIN spelling hands the reftype over as a bare `$id` rather than
        // a `ref(…)` call. Returning "" here meant `ref.get_desc` was resolved
        // against no type at all for `br_on_cast_desc_eq $l anyref $a`.
        Some(ExprKind::Ident(n)) => n.clone(),
        _ => String::new(),
    }
}

/// `br_on_cast_desc_eq $l rt1 rt2` / `br_on_cast_desc_eq_fail` — the Custom
/// Descriptors branching casts.
///
/// Stack operands are the reference and, ON TOP OF IT, the descriptor to
/// compare against: the proposal puts the descriptor last, the same as
/// `struct.new_desc`. `args` are the immediates in source order — label,
/// source reftype, target reftype.
///
/// The semantics are pinned by the proposal's own suite,
/// `test/core/custom-descriptors/br_on_cast_desc_eq.wast`:
///
///   * A NULL DESCRIPTOR TRAPS, and it traps FIRST — `self-nullable-null-null`
///     passes a null reference AND a null descriptor and is asserted to trap,
///     not to branch. So the descriptor is checked before the reference is
///     looked at at all.
///   * A null reference matches iff the TARGET reftype is nullable:
///     `…-nullable-null-desc` → 1 for `(ref null $a)`, `…-nonnullable-null-desc`
///     → 0 for `(ref $a)`. This is why the nullability of the immediate has to
///     survive lowering rather than being folded away to a heap type.
///   * Otherwise the test is descriptor IDENTITY, not structural equality:
///     `…-val-other` holds a second descriptor of the same type with equal
///     contents and is asserted 0. `ref_eq` is the right comparison; a value
///     compare would answer 1.
///
/// `ref.get_desc` traps on a null reference (its result type is non-nullable),
/// so the null-reference answer is taken from the immediate and never reaches
/// it — reordering these two branches would turn a legal `…-null-desc` case
/// into a trap.
fn emit_br_on_cast_desc_eq_stmt(
    __w: &mut WastWalker,
    args: &[Expression],
    is_fail: bool,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    let target = br_target_of(args.first());
    let to_rt = args.get(2);
    let target_nullable = reftype_is_nullable(to_rt);
    let to_name = reftype_heap_name(to_rt);

    // ── The real instruction, whenever the target is a block ──────────────
    //
    // Everything below this point is a LOWERING — `ref.is_null` + `if`/`else`
    // computing the match by hand. It exists because the VM's own
    // `BR_ON_CAST_DESC_EQ` used to answer a null reference wrongly (it read
    // `matched = !val.is_null_ref() && …`, so a null never matched even a
    // NULLABLE target). That is fixed, so the instruction can be emitted as
    // itself and the lowering kept only for the shapes it still covers:
    //
    //   * a LOOP target, where a branch is a `continue` carrying the loop's
    //     params rather than a block's results, and
    //   * the function-level implicit label, where a branch is a `return`.
    //
    // ⚠ THE LANDING PAD IS NOT DECORATION. This walker models the wasm
    // operand stack as AST TEMPORARIES, so a block's results are read out of
    // `result_temps` — but a VM-level branch jumps without ever running the
    // assignments that fill them. Branching first to a private inner block
    // gives the taken path somewhere to write the temps before it continues
    // to the real target. The wrapper is ordinary structured control flow;
    // the descriptor instruction itself is emitted exactly as spelled.
    if let Some(entry) = labels.resolve(&target) {
        if matches!(entry.kind, LabelKind::Block) {
            emit_br_on_cast_desc_eq_native(
                __w, args, is_fail, span, &entry, statements, stack,
            );
            return;
        }
    }

    // The descriptor is the TOPMOST operand, so it pops first.
    let desc_val = stack.pop().unwrap_or_else(Expression::null);
    let ref_val = stack.pop().unwrap_or_else(Expression::null);
    let dtmp = fresh_result_temp(__w);
    let rtmp = fresh_result_temp(__w);
    for (name, init) in [(&dtmp, desc_val), (&rtmp, ref_val)] {
        statements.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name.clone()),
                type_hint: None,
                init: Some(init),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    }

    // 1. Null descriptor → trap, whatever the reference is.
    //
    // Routed through `ref.cast_desc_eq` rather than a bare `unreachable`.
    // The VM already traps "null descriptor reference" for a null descriptor,
    // unconditionally and before the reference is looked at
    // (`dispatch.rs`, `Op::REF_CAST_DESC_EQ`), which is exactly this case —
    // and it is the MESSAGE the spec fixtures assert. `unreachable` trapped
    // too, so the semantics were right, but it reported "unreachable
    // executed"; that went unnoticed for as long as `assert_trap` parsed the
    // expected message and threw it away.
    let null_desc_trap = match to_rt {
        Some(rt) => Statement::with_span(
            StmtKind::Expr(make_call(
                "ref_cast_desc_eq",
                vec![
                    rt.clone(),
                    Expression::ident(&rtmp),
                    Expression::ident(&dtmp),
                ],
                span,
            )),
            span,
        ),
        None => Statement::with_span(StmtKind::Expr(trap_expr()), span),
    };
    statements.push(Statement::with_span(
        StmtKind::If {
            cond: make_call("ref_is_null", vec![Expression::ident(&dtmp)], span),
            then_body: vec![null_desc_trap],
            else_body: None,
            elifs: Vec::new(),
        },
        span,
    ));

    // 2. `matched` starts at the answer for a NULL reference, which the target
    //    reftype's nullability decides on its own.
    let mtmp = fresh_result_temp(__w);
    statements.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(mtmp.clone()),
            type_hint: None,
            init: Some(Expression::int(i64::from(target_nullable))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));

    // 3. Non-null reference → compare descriptor IDENTITY. Guarded, so
    //    `ref.get_desc` never sees the null that would trap it.
    let is_non_null = Expression::new(ExprKind::Binary {
        op: BinOp::StrictEq,
        left: Box::new(make_call(
            "ref_is_null",
            vec![Expression::ident(&rtmp)],
            span,
        )),
        right: Box::new(Expression::int(0)),
    });
    let got_desc = make_call(
        "__wast_ref_get_desc",
        vec![Expression::ident(&rtmp), Expression::string(&to_name)],
        span,
    );
    statements.push(Statement::with_span(
        StmtKind::If {
            cond: is_non_null,
            then_body: vec![Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::ident(&mtmp)),
                    value: Box::new(make_call(
                        "ref_eq",
                        vec![got_desc, Expression::ident(&dtmp)],
                        span,
                    )),
                },
            )))],
            else_body: None,
            elifs: Vec::new(),
        },
        span,
    ));

    let cond = if is_fail {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::ident(&mtmp)),
            right: Box::new(Expression::int(0)),
        })
    } else {
        Expression::ident(&mtmp)
    };
    let mut then_body: Vec<Statement> = Vec::new();
    match labels.resolve_dest(&target) {
        BrDest::Frame(entry) => {
            // ⚠ THE BRANCH CARRIES EVERY RESULT, NOT JUST THE REFERENCE.
            //
            // The typing rule is `t* rt_1 (ref null (exact y)) -> t* (rt_1 \ rt_2)`
            // — that leading `t*` is real. The proposal's "Sent values" section
            // exercises it directly:
            //
            //   (block (result i32 i32 eqref)
            //     (i32.const 1) (i32.const 2) (ref.null none)
            //     (br_on_cast_desc_eq 0 eqref (ref null $a) (global.get $b1)))
            //
            // Assigning only the topmost result left the two i32s unwritten, so
            // `cast-succeeds-null` read garbage and trapped in its own
            // assert helper. The reference is the TOP result on both forms
            // (`_fail` branches with `rt_1 \ rt_2`, still the reference); the
            // values below it come off the stack, untouched and NOT consumed —
            // the fall-through path still needs them.
            let temps = &entry.result_temps;
            if let Some((ref_temp, below)) = temps.split_last() {
                carry_stack_into_temps(below, stack, false, &mut then_body);
                then_body.push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Assign {
                        target: Box::new(Expression::ident(ref_temp)),
                        value: Box::new(Expression::ident(&rtmp)),
                    },
                ))));
            }
            then_body.push(br_stmt_for(&entry, span));
        }
        // ⚠ THE FUNCTION-LEVEL LABEL IS A RETURN, NOT A BARE BREAK.
        //
        // `br_on_cast 0` written directly in a function body targets the
        // function's implicit block, and branching to it RETURNS carrying the
        // reference. Emitting `Break(Implicit)` left the value behind and fell
        // through to whatever followed — in the spec fixtures, an
        // `(unreachable)` placed there precisely because the branch should have
        // jumped over it. The same instruction inside an explicit `(block …)`
        // resolved to a real label and worked, so the two spellings disagreed.
        //
        // ⛔ AND IT CARRIES `t*` TOO, not just the reference. Returning the
        // ref alone is right only for a single-result function; the typing
        // rule's leading `t*` is as real here as it is on the block path
        // directly above, where assigning only the topmost result delivered
        // garbage for the values underneath.
        BrDest::Func => {
            let n = __w.current_fn_results;
            let below = n.saturating_sub(1).min(stack.len());
            let mut vals: Vec<Expression> = stack[stack.len() - below..].to_vec();
            vals.push(Expression::ident(&rtmp));
            then_body.push(func_return_stmt_n(n, &mut vals, true, span));
        }
        BrDest::Unresolved => then_body.push(Statement::with_span(
            StmtKind::Return(Some(Expression::ident(&rtmp))),
            span,
        )),
    }
    statements.push(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body: None,
            elifs: Vec::new(),
        },
        span,
    ));
    // Fall-through: the reference stays for the continuation.
    stack.push(Expression::ident(&rtmp));
}

/// `br_on_cast_desc_eq` / `_fail` emitted as the real instruction.
///
/// Shape, in wasm terms:
///
/// ```text
/// block $outer                       ;; void
///   block $inner                     ;; void
///     local.get $ref  local.get $desc
///     br_on_cast_desc_eq $inner ht_1 ht_2   ;; matched -> leaves $inner
///     br $outer                              ;; no match -> skip the pad
///   end
///   ;; landing pad: fill the target block's result temps, then branch on
///   <result temps> = <stack values> , <ref>
///   br $target
/// end
/// ```
///
/// The null-descriptor trap is NOT emitted here: the instruction traps on it
/// itself, before it looks at the reference, and with the message the suite
/// asserts (`"null descriptor reference"`). The lowering had to route that
/// through `ref.cast_desc_eq` to borrow the wording.
fn emit_br_on_cast_desc_eq_native(
    __w: &mut WastWalker,
    args: &[Expression],
    is_fail: bool,
    span: Span,
    entry: &LabelEntry,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    // The descriptor is the TOPMOST operand, so it pops first. Both are
    // materialised into locals before the branch: the landing pad needs the
    // reference AFTER a branch has already discarded the operand stack, and
    // a local is the only place it survives.
    let desc_val = stack.pop().unwrap_or_else(Expression::null);
    let ref_val = stack.pop().unwrap_or_else(Expression::null);
    let dtmp = fresh_result_temp(__w);
    let rtmp = fresh_result_temp(__w);
    for (name, init) in [(&dtmp, desc_val), (&rtmp, ref_val)] {
        statements.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name.clone()),
                type_hint: None,
                init: Some(init),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    }

    let outer = fresh_block_label(__w);
    let inner = fresh_block_label(__w);
    let head = if is_fail {
        "br_on_cast_desc_eq_fail"
    } else {
        "br_on_cast_desc_eq"
    };
    // Immediates in the order the compiler reads them: label, ht_1 (source),
    // ht_2 (target); then the two stack operands.
    let instr = make_call(
        head,
        vec![
            Expression::string(&inner),
            args.get(1).cloned().unwrap_or_else(Expression::null),
            args.get(2).cloned().unwrap_or_else(Expression::null),
            Expression::ident(&rtmp),
            Expression::ident(&dtmp),
        ],
        span,
    );
    let inner_body = vec![
        Statement::with_span(StmtKind::Expr(instr), span),
        Statement::with_span(
            StmtKind::Break(BreakTarget::Label(outer.clone())),
            span,
        ),
    ];
    let mut outer_body = vec![Statement::with_span(
        StmtKind::Labeled {
            label: inner,
            body: Box::new(Statement::with_span(StmtKind::Block(inner_body), span)),
        },
        span,
    )];
    // The landing pad. The branch carries every result, not just the
    // reference: the typing rule's leading `t*` is real, and the values below
    // the reference are NOT consumed because the fall-through still needs
    // them.
    if let Some((ref_temp, below)) = entry.result_temps.split_last() {
        carry_stack_into_temps(below, stack, false, &mut outer_body);
        outer_body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Assign {
                target: Box::new(Expression::ident(ref_temp)),
                value: Box::new(Expression::ident(&rtmp)),
            },
        ))));
    }
    outer_body.push(br_stmt_for(entry, span));
    statements.push(Statement::with_span(
        StmtKind::Labeled {
            label: outer,
            body: Box::new(Statement::with_span(StmtKind::Block(outer_body), span)),
        },
        span,
    ));
    // Fall-through: the reference stays for the continuation.
    stack.push(Expression::ident(&rtmp));
}

fn emit_folded_stmtwise(__w: &mut WastWalker,
    inner: Pair<Rule>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) -> Result<(), String> {
    let head = folded_instr_head(&inner);
    let mut immediate_args: Vec<Expression> = Vec::new();
    for child in inner.clone().into_inner() {
        match child.as_rule() {
            Rule::instr_arg => {
                let nested_folded = child
                    .clone()
                    .into_inner()
                    .next()
                    .filter(|x| x.as_rule() == Rule::folded_instr);
                match nested_folded {
                    // `push_folded_operand` already routes stmt-lowering cases
                    // onward AND implements the spec's partial-fold rule (an
                    // operand may nest fewer values than it consumes, taking
                    // the rest from the enclosing stack). Calling
                    // `walk_folded_instr_as_expr` directly skipped that, so
                    // `(return (i32.add (i32.const 100)))` after a block that
                    // left 7 on the stack answered 100 instead of 107 — the
                    // unfolded `(i32.add (i32.const 100)) (return)` was right.
                    Some(op) => {
                        push_folded_operand(__w, op, labels, statements, stack)?;
                    }
                    None => immediate_args.push(walk_instr_arg_for(__w, child, labels, &head)?),
                }
            }
            Rule::instr => {
                let nested = child.clone().into_inner().next();
                match nested {
                    Some(op)
                        if op.as_rule() == Rule::folded_instr
                            && folded_needs_stmt_lowering(&op) =>
                    {
                        emit_folded_operand_stmtwise(__w, op, labels, statements, stack)?;
                    }
                    _ => stack.push(walk_instr_as_expr(__w, child, labels)?),
                }
            }
            _ => {}
        }
    }
    match head.as_str() {
        "br" => {
            emit_br_stmt_carry(__w, immediate_args.first(), span, labels, statements, stack);
        }
        // Tail calls, same lowering as the plain arm: qualify the callee
        // through `call`/`call_ref`, then hand it to `__wasm_return_call` so
        // the compiler emits the frame-reusing `Op::RETURN_CALL`. Without this
        // the folded form fell to the expression walk, which produced a plain
        // call to an UNQUALIFIED callee — "null is not callable".
        "return_call" | "return_call_ref" => {
            // ⛔ `return_call_ref` does NOT share `return_call`'s lowering.
            // `__wasm_return_call` takes a QUALIFIED CALLEE NAME — it is how a
            // `return_call $f` names the module method to tail-call — while
            // `return_call_ref` has a funcref VALUE and its own opcode. Feeding
            // the value in where the name belongs is what made
            // `return_call_ref.wast` report "null is not callable".
            let inner_name = if head == "return_call" {
                "call"
            } else {
                "return_call_ref"
            };
            let arity = get_instruction_arity(__w, inner_name, &immediate_args);
            let mut args = immediate_args;
            let pop_count = arity.min(stack.len());
            let drain_start = stack.len() - pop_count;
            let popped: Vec<Expression> = stack.drain(drain_start..).collect();
            args.extend(popped);
            let call = map_instr_to_ast(__w, inner_name.to_string(), args, span)?;
            if inner_name == "return_call_ref" {
                statements.push(Statement::with_span(StmtKind::Expr(call), span));
            } else if let ExprKind::Call {
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
                statements.push(Statement::with_span(StmtKind::Return(Some(call)), span));
            }
        }
        // Folding is sugar for the flat sequence, so these read their tag
        // immediate off the PAIR and their payload off the stack the nested
        // operands just evaluated onto — identical to the plain arms.
        "throw" => {
            let tag = peek_instr_tag_ref(__w, &inner).unwrap_or_default();
            let arity = tag_arity(__w, &tag) as usize;
            let n = arity.min(stack.len());
            let args: Vec<Expression> = stack.split_off(stack.len() - n);
            statements.push(Statement::with_span(
                StmtKind::WasmThrow { tag, args },
                span,
            ));
        }
        "throw_ref" => {
            let exnref_expr = stack.pop().unwrap_or_else(Expression::null);
            let exnref_local = match &exnref_expr.kind {
                ExprKind::Ident(n) => n.clone(),
                _ => {
                    let tmp = fresh_result_temp(__w);
                    statements.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(tmp.clone()),
                            type_hint: None,
                            init: Some(exnref_expr),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }));
                    tmp
                }
            };
            statements.push(Statement::with_span(
                StmtKind::WasmRethrow { exnref_local },
                span,
            ));
        }
        "br_if" => {
            let cond = stack.pop().unwrap_or(Expression::int(0));
            emit_br_if_stmt(__w, 
                immediate_args.first(),
                cond,
                span,
                labels,
                statements,
                stack,
            );
        }
        "br_table" => {
            emit_br_table_stmt(__w, &immediate_args, span, labels, statements, stack);
        }
        "br_on_cast" | "br_on_cast_fail" => {
            emit_br_on_cast_stmt(__w,
                &immediate_args,
                head == "br_on_cast_fail",
                span,
                labels,
                statements,
                stack,
            );
        }
        "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail" => {
            emit_br_on_cast_desc_eq_stmt(
                __w,
                &immediate_args,
                head == "br_on_cast_desc_eq_fail",
                span,
                labels,
                statements,
                stack,
            );
        }
        "return" => {
            let n = __w.current_fn_results;
            if n >= 2 {
                statements.push(multi_value_return_stmt(stack, n, span));
            } else {
                let val = stack.pop();
                statements.push(Statement::with_span(StmtKind::Return(val), span));
            }
        }
        // WASM `select` pops (val1, val2, cond) and evaluates all three — no
        // short-circuit. Bind each to a temp in program order (they may be
        // DEFERRED side-effectful expressions), then pick via a ternary over
        // the temps.
        "select" => {
            let n = 3.min(stack.len());
            let mut ops: Vec<Expression> = stack.split_off(stack.len() - n);
            while ops.len() < 3 {
                ops.insert(0, Expression::null());
            }
            let cond = ops.pop().expect("three operands");
            let val2 = ops.pop().expect("three operands");
            let val1 = ops.pop().expect("three operands");
            let mut bind = |e: Expression| -> Expression {
                let t = fresh_result_temp(__w);
                statements.push(Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(t.clone()),
                        type_hint: None,
                        init: Some(e),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }));
                Expression::ident(&t)
            };
            let v1 = bind(val1);
            let v2 = bind(val2);
            let c = bind(cond);
            stack.push(Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(c),
                    then: Box::new(v1),
                    else_: Box::new(v2),
                },
                span,
            ));
        }
        // Mirrors the plain arm: argc/results from the `(type $sig)` type-use,
        // operands (spec order: call args, then the table index on top) from
        // the stack the nested operands just evaluated onto.
        "call_indirect" | "return_call_indirect" => {
            let (argc, expected_results) = peek_typeuse_shape(__w, &inner);
            let tableidx = peek_call_indirect_table(&inner)
                .map(|t| resolve_table_index(__w, &t) as usize)
                .unwrap_or_else(|| __w.table_index_base);
            let n = (argc + 1).min(stack.len());
            let operands: Vec<Expression> = stack.split_off(stack.len() - n);
            // The DECLARED functype rides along as a fourth immediate. The
            // instruction's own three are counts, and a count cannot pick
            // between two same-arity types — see `Chunk::call_indirect_sigs`.
            // `U8_U8_U8` is used by exactly two opcodes, both of them this one,
            // so widening the arg list here reaches nothing else.
            let signature = typeuse_signature(__w, &inner);
            let canon = typeuse_canon_name(__w, &inner);
            let mut call_args = vec![
                Expression::int(argc as i64),
                Expression::int(tableidx as i64),
                Expression::int(expected_results as i64),
                Expression::string(&signature),
                Expression::string(&canon),
            ];
            call_args.extend(operands);
            if head == "return_call_indirect" {
                let call = make_call("return_call_indirect", call_args, span);
                statements.push(Statement::with_span(StmtKind::Expr(call), span));
            } else {
                // ⚠ NOT `land_instr_value`. A direct `call` to a multi-result
                // function destructures into N temps there, and doing the same
                // here looks right but is not: the compiler only recognises the
                // multi-value shape for a DIRECT call, so the destructure reads
                // the packed result as an array and every consumer gets the
                // wrong values. Returning a pair out of `(result i32 i32)` needs
                // that compiler half; until it exists the packed push is the
                // shape everything else in the pipeline expects.
                stack.push(make_call("call_indirect", call_args, span));
            }
        }
        _ => {
            // Ordinary head: apply it to the stack exactly like the flat form.
            let arity = get_instruction_arity(__w, &head, &immediate_args);
            let pop_count = arity.min(stack.len());
            let drain_start = stack.len() - pop_count;
            let popped: Vec<Expression> = stack.drain(drain_start..).collect();
            let mut args = immediate_args;
            args.extend(popped);
            // Same gap in the flat form.
            let call_pushes = if head == "call" {
                call_result_count(__w, &args)
            } else {
                1
            };
            let expr = map_instr_to_ast(__w, head.clone(), args, span)?;
            if call_pushes >= 2 {
                land_instr_value(__w, expr, call_pushes, true, span, statements, stack);
            } else if get_instruction_push_count(&head) > 0 {
                stack.push(expr);
            } else {
                statements.push(Statement::with_span(StmtKind::Expr(expr), span));
            }
        }
    }
    Ok(())
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
    param_temps: Vec<String>,
}

/// A fresh synthetic block/loop label.
fn fresh_block_label(__w: &mut WastWalker) -> String {
    let name = format!("__wat_lbl{}", __w.wast_temp_counter);
    __w.wast_temp_counter += 1;
    name
}

struct LabelStack(Vec<LabelEntry>);

impl LabelStack {
    fn new() -> Self {
        LabelStack(Vec::new())
    }
    /// How many labels are in scope. A branch to depth `len()` names the
    /// function's own implicit label — the one a `return` targets.
    fn len(&self) -> usize {
        self.0.len()
    }
    /// Push a frame, minting a synthetic name if the source has none. Returns the
    /// effective label so the caller can build the matching `Labeled` statement.
    fn push(
        &mut self,
        __w: &mut WastWalker,
        name: Option<String>,
        kind: LabelKind,
        result_temps: Vec<String>,
    ) -> String {
        let effective = match name {
            Some(n) => n,
            None => fresh_block_label(__w),
        };
        self.0.push(LabelEntry {
            name: effective.clone(),
            kind,
            result_temps,
            param_temps: Vec::new(),
        });
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
            BrTarget::Innermost => self.0.last().cloned(),
        }
    }

    /// Resolve a branch target to where it LANDS, distinguishing the function's
    /// own implicit label from a target that is genuinely out of scope. See
    /// `BrDest` for why the two must not share an answer.
    fn resolve_dest(&self, target: &BrTarget) -> BrDest {
        if let Some(entry) = self.resolve(target) {
            return BrDest::Frame(entry);
        }
        let is_func_label = match target {
            BrTarget::Index(d) => *d == self.0.len(),
            // `br` with no label immediate inside a function with no enclosing
            // block names the function label by the same reasoning.
            BrTarget::Innermost => self.0.is_empty(),
            BrTarget::Named(_) => false,
        };
        if is_func_label {
            BrDest::Func
        } else {
            BrDest::Unresolved
        }
    }
}

/// The `return` a branch to the function's implicit label lowers to: the top
/// `n` stack values become the function's results.
///
/// `n == 0` is NOT a fall-through — it is a bare `return` that DISCARDS
/// whatever the branch unwinds past. `unwind.wast`'s `func-unwind-by-br_table`
/// is `(i32.const 3) (i64.const 1) (br_table 0 (i32.const 0))` in a
/// result-less function: carrying values alone leaves that one failing.
///
/// `consume` mirrors `carry_stack_into_temps`: a `br` owns the stack it
/// unwinds, but a `br_if`/`br_table` case must leave it intact for the paths
/// that are not taken.
fn func_return_stmt_n(
    n: usize,
    stack: &mut Vec<Expression>,
    consume: bool,
    span: Span,
) -> Statement {
    if n == 0 {
        return Statement::with_span(StmtKind::Return(None), span);
    }
    let avail = n.min(stack.len());
    let start = stack.len() - avail;
    let vals: Vec<Expression> = if consume {
        stack.split_off(start)
    } else {
        stack[start..].to_vec()
    };
    let val = if n >= 2 {
        Some(Expression::new(ExprKind::Tuple(vals)))
    } else {
        vals.into_iter().next()
    };
    Statement::with_span(StmtKind::Return(val), span)
}

/// How a `br`/`br_if` names its destination frame.
enum BrTarget {
    Named(String),
    Index(usize),
    Innermost,
}

/// Where a branch actually lands.
///
/// ⛔ `LabelStack::resolve` answers `None` for TWO different things: a target
/// that is out of scope, and the function's OWN implicit label at depth
/// `len()`. Every branch emitter read that single `None` as
/// `Break(BreakTarget::Implicit)` — which at function top level breaks nothing
/// and FALLS THROUGH. Per §4.4.8 a branch to the outermost label is a `return`
/// carrying the function's results, so `(br 0 (i32.const 50)) (i32.const 51)`
/// returned 51. `return` was right, `br` to a block was right, and only the
/// sibling spelling was wrong — invisible to validation, since the module is
/// perfectly valid and the failure is a WRONG VALUE.
///
/// Matching on this enum instead of on `Option` is what makes the next branch
/// instruction unable to forget: an unhandled variant is a compile error.
enum BrDest {
    Frame(LabelEntry),
    /// The function's implicit label — a branch here IS a return.
    Func,
    /// Neither in scope nor the function label: an invalid module.
    ///
    /// ⛔ NOT AN `Option`. The first cut of this returned `Option<BrDest>`,
    /// and `None` is exactly the shape that let the bug spread in the first
    /// place — a site can handle it without ever deciding what the function
    /// label means, and three of the six branch emitters did precisely that.
    /// A third variant makes every `match` exhaustive, so forgetting one is a
    /// compile error instead of a spelling that silently falls through.
    Unresolved,
}

/// Derive a `br` target from its first argument (label id or numeric index).
fn br_target_of(arg: Option<&Expression>) -> BrTarget {
    match arg.map(|a| &a.kind) {
        Some(ExprKind::Ident(n)) => BrTarget::Named(n.clone()),
        Some(ExprKind::Lit(Literal::Int(i))) => BrTarget::Index(*i as usize),
        _ => BrTarget::Innermost,
    }
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
        ),
    }
}

/// Unconditional `br`: carry (consume) the top N stack values into the target's
/// temps, then jump. A `br` to a loop is a continue carrying the NEXT
/// iteration's params; a `br` to a block carries the block's results.
fn emit_br_stmt_carry(__w: &mut WastWalker, 
    lbl_arg: Option<&Expression>,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    // Deferred stack values (calls, loads) must evaluate in program order —
    // BEFORE the branch — exactly once, including the ones the unwinding
    // discards.
    preserve_stack_across_block(__w, stack, statements);
    let target = br_target_of(lbl_arg);
    match labels.resolve_dest(&target) {
        BrDest::Frame(entry) => {
            let carry = branch_carry_temps(&entry);
            carry_stack_into_temps(&carry, stack, true, statements);
            statements.push(br_stmt_for(&entry, span));
        }
        BrDest::Func => {
            let n = __w.current_fn_results;
            statements.push(func_return_stmt_n(n, stack, true, span));
        }
        BrDest::Unresolved => statements.push(make_br_stmt_opt(None, labels, span)),
    }
}

/// Conditional `br_if`: the carried values pass through only when taken, so
/// peek (don't consume) the top N into the target's temps inside the then-arm.
fn emit_br_if_stmt(__w: &mut WastWalker, 
    lbl_arg: Option<&Expression>,
    cond: Expression,
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    // Spec order: the carried value(s) were pushed BEFORE the condition, so
    // any deferred stack expressions must evaluate first, exactly once —
    // taken or not. The condition then evaluates inside the `if` header,
    // after these materialized bindings.
    preserve_stack_across_block(__w, stack, statements);
    let target = br_target_of(lbl_arg);
    let mut then_body: Vec<Statement> = Vec::new();
    let branch = match labels.resolve_dest(&target) {
        BrDest::Frame(entry) => {
            let carry = branch_carry_temps(&entry);
            carry_stack_into_temps(&carry, stack, false, &mut then_body);
            br_stmt_for(&entry, span)
        }
        // ⛔ Peek, never consume: the untaken path still owns this stack.
        BrDest::Func => {
            let n = __w.current_fn_results;
            func_return_stmt_n(n, stack, false, span)
        }
        BrDest::Unresolved => make_br_stmt_opt(None, labels, span),
    };
    then_body.push(branch);
    statements.push(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            else_body: None,
            elifs: Vec::new(),
        },
        span,
    ));
}

/// `br_table l0 l1 … ln`: pops a selector index and branches to the l_index
/// frame (l_n is the default). Lowered to an if/else-if chain over the index
/// bound to a temp; each case carries the same top-N snapshot into the chosen
/// target's result temps before branching, mirroring `br`.
fn emit_br_table_stmt(__w: &mut WastWalker, 
    target_args: &[Expression],
    span: Span,
    labels: &mut LabelStack,
    statements: &mut Vec<Statement>,
    stack: &mut Vec<Expression>,
) {
    let targets: Vec<BrTarget> = target_args.iter().map(|a| br_target_of(Some(a))).collect();
    let index = stack.pop().unwrap_or(Expression::int(0));
    // The carried values were pushed BEFORE the selector: materialize any
    // deferred expressions so they evaluate first and exactly once (each case
    // below references the same snapshot; only the taken case runs, but the
    // VALUES must have been computed regardless).
    preserve_stack_across_block(__w, stack, statements);
    let carried: Vec<Expression> = stack.clone();
    let fn_results = __w.current_fn_results;
    let br_for = |t: &BrTarget| -> Vec<Statement> {
        match labels.resolve_dest(t) {
            BrDest::Func => {
                // Each case reads the same snapshot; only one runs, so this
                // takes a private copy rather than draining `carried`.
                let mut snap = carried.clone();
                vec![func_return_stmt_n(fn_results, &mut snap, true, span)]
            }
            BrDest::Frame(entry) => {
                let mut out = Vec::new();
                let n = entry.result_temps.len();
                let start = carried.len().saturating_sub(n);
                for (k, tmp) in entry.result_temps.iter().enumerate() {
                    if let Some(val) = carried.get(start + k) {
                        out.push(Statement::new(StmtKind::Expr(Expression::new(
                            ExprKind::Assign {
                                target: Box::new(Expression::ident(tmp)),
                                value: Box::new(val.clone()),
                            },
                        ))));
                    }
                }
                out.push(br_stmt_for(&entry, span));
                out
            }
            BrDest::Unresolved => vec![make_br_stmt_opt(None, labels, span)],
        }
    };
    if targets.is_empty() {
        // Degenerate: nothing to branch to.
    } else if targets.len() == 1 {
        // The selector still evaluates (exactly once) even though only one
        // target exists.
        statements.push(Statement::with_span(StmtKind::Expr(index), span));
        statements.extend(br_for(&targets[0]));
    } else {
        let idx_tmp = fresh_result_temp(__w);
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
        // ⛔ NEITHER A NESTED CHAIN NOR `elifs` — A FLAT SEQUENCE. Both of
        // those nest each successive target one level deeper, and a branch
        // DEPTH is a u8: `br_table.wast`'s "large" case has 16150 targets, so
        // every index past 255 branched to the wrong block. (The nested form
        // also cost one Rust frame per target in the compiler and aborted the
        // process long before that; the depth bug was hiding behind it.)
        //
        // Emitting the tests as INDEPENDENT `if`s at the SAME level keeps every
        // branch depth constant. Only one can run: each body BRANCHES, so
        // control leaves before the next test is reached, and falling past all
        // of them is exactly the default case.
        for k in 0..targets.len() - 1 {
            statements.push(Statement::with_span(
                StmtKind::If {
                    cond: Expression::new(ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(Expression::ident(&idx_tmp)),
                        right: Box::new(Expression::int(k as i64)),
                    }),
                    then_body: br_for(&targets[k]),
                    elifs: Vec::new(),
                    else_body: None,
                },
                span,
            ));
        }
        // The last label is `br_table`'s default: reached only by falling past
        // every test above.
        statements.extend(br_for(&targets[targets.len() - 1]));
    }
}

// ── Entry point ──────────────────────────────────────────────────────────────

/// The `spectest` module every wast script may import from. The reference
/// interpreter predefines it (`interpreter/host/spectest.ml`,
/// `interpreter/README.md` §"Spectest host module"), and 15 of the suite's
/// files import from it — declaring an import of it always worked, but CALLING
/// one was an "Unresolved import" because nothing implemented it.
///
/// ⛔ It is a synthesized MODULE, not a set of host functions. The spec defines
/// it as a module type, the script runner's own `(register …)` machinery
/// already resolves module-to-module imports, and a host function would put
/// test-harness scaffolding into the runtime's product surface for no gain.
///
/// The `print*` functions are observable only as stdout in the reference
/// interpreter; no assertion in the suite reads that, so the bodies are empty
/// and what matters is that a call SUCCEEDS. The global values are the
/// interpreter's exactly — 666 for the integers, **666.6** for the floats
/// (`Num (F32 (F32.of_float 666.6))`), which `global.wast` reads back.
const SPECTEST_MODULE: &str = r#"
(module
  (global (export "global_i32") i32 (i32.const 666))
  (global (export "global_i64") i64 (i64.const 666))
  (global (export "global_f32") f32 (f32.const 666.6))
  (global (export "global_f64") f64 (f64.const 666.6))
  (table (export "table") 10 20 funcref)
  (table (export "table64") i64 10 20 funcref)
  (memory (export "memory") 1 2)
  (func (export "print"))
  (func (export "print_i32") (param i32))
  (func (export "print_i64") (param i64))
  (func (export "print_f32") (param f32))
  (func (export "print_f64") (param f64))
  (func (export "print_i32_f32") (param i32 f32))
  (func (export "print_f64_f64") (param f64 f64))
)
(register "spectest")
"#;

/// Walk `SPECTEST_MODULE` into `body` ahead of the script's own commands, so
/// its module class exists and is registered before anything can import it.
fn prepend_spectest_module(__w: &mut WastWalker, body: &mut Vec<Statement>) -> Result<(), String> {
    let pairs = WastParser::parse(Rule::program, SPECTEST_MODULE)
        .map_err(|e| format!("internal: the spectest module does not parse: {e}"))?;
    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for cmd in top.into_inner() {
                    if cmd.as_rule() != Rule::EOI {
                        walk_script_cmd(__w, cmd, body)?;
                    }
                }
            }
            Rule::EOI => {}
            _ => walk_script_cmd(__w, top, body)?,
        }
    }
    Ok(())
}

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs =
        WastParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;

    // Every registry this walk keeps, created here and dropped when `parse`
    // returns — including on the `?` paths below. A wast SCRIPT accumulates
    // across its own modules (index bases, module numbering, registered
    // exports); nothing accumulates across scripts.
    let mut __w_owned = WastWalker::default();
    let __w = &mut __w_owned;
    let mut body = Vec::new();
    // The `spectest` host module the script runner predefines — ONLY when this
    // script actually names it.
    //
    // ⛔ It is a real module, so it TAKES INDEX SPACE: its `(memory …)` becomes
    // memory 0 and every later module's base shifts by one. Prepending it
    // unconditionally moved slot 0 out from under every script that never
    // mentions spectest — 48 corpus tests (all the v128 load/store and
    // stringref files, which read and write memory 0) went red at once. The
    // 15 files that import it get it; nothing else is touched.
    if source.contains("spectest") {
        prepend_spectest_module(__w, &mut body)?;
    }
    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for cmd in top.into_inner() {
                    if cmd.as_rule() != Rule::EOI {
                        walk_script_cmd(__w, cmd, &mut body)?;
                    }
                }
            }
            Rule::EOI => {}
            _ => walk_script_cmd(__w, top, &mut body)?,
        }
    }

    // Prepend the `assert_trap` helper when this script has any. Prepended,
    // not appended: the asserts that call it run at top level in source order.
    if __w_owned.needs_trap_contains {
        body.insert(0, build_trap_check_helper(Span::default()));
    }

    Ok(Module {
        // The canon section this walk built, in canonidx order. Empty unless
        // the source declared a `(component …)`.
        canon: vybe_ast::canon::ComponentSection {
            defs: std::mem::take(&mut __w_owned.canon_section),
            types: std::mem::take(&mut __w_owned.comp_types),
            funcs: std::mem::take(&mut __w_owned.comp_func_space)
                .into_iter()
                .map(|f| f.canonidx)
                .collect(),
        },
        name: "main".into(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
        // ⛔ A WAST FUNCTION IS NOT A JAVASCRIPT FUNCTION OBJECT.
        //
        // The walker expresses a `(module …)` as a `ClassDecl` whose members
        // are its funcs — a container, so that a script's several modules keep
        // separate index spaces (`__wasm_module`, `__wasm_module_1`, …) and
        // passive elem segments can resolve `ref.func` through the right one.
        //
        // Shared class emission then did the only thing a class full of methods
        // means in the ECMA object model: stamped `name` / `length` /
        // `prototype` and a `__nonenum` set onto every one of them. In a
        // module that declares no such fields, that is a `struct.set` against
        // nothing — and it is why NO wast module we emitted would load on a
        // spec engine. wast maps to INSTRUCTIONS; it has no function objects.
        directives: vybe_ast::Directives {
            functions_are_objects: Some(false),
            ..Default::default()
        },
    })
}

// ── Script commands ───────────────────────────────────────────────────────────

fn walk_script_cmd(__w: &mut WastWalker, pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::script_cmd => {
            let inner = pair.into_inner().next().ok_or("Empty script_cmd")?;
            walk_script_cmd(__w, inner, body)
        }
        // `inline_module` is a module's field list written bare at the top level
        // — structurally identical to `module` minus the wrapper, so it walks
        // through the same path (it has no `id`, hence no named-module target).
        Rule::module | Rule::inline_module => {
            body.extend(walk_module(__w, pair)?);
            Ok(())
        }
        // `(component …)` — the Component Model's OUTER text format
        // (`Explainer.md` §2). Not a module field and not core WASM: it WRAPS
        // core modules and adds the spaces core WASM has no syntax for.
        Rule::component => walk_component(__w, pair, body).map(|_| ()),
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
        // `(module binary "…")` embeds a module as raw BYTES.
        //
        // This used to be `Ok(())` — accepted and skipped. A top-level binary
        // module in a spec fixture is an assertion in its own right: it is
        // stated to be WELL-FORMED, and the file is testing that an
        // implementation accepts it. Skipping it means the one thing the
        // fixture asserts goes unchecked, and it also hides the opposite
        // failure — an over-strict decoder rejecting a module the spec says is
        // fine, which is exactly the risk a new validation pass introduces.
        //
        // So: decode it, and report a failure to decode. Not yet
        // INSTANTIATED — its exports are not reachable by a later `invoke`,
        // which needs the decoded chunks spliced into the script and is a
        // separate piece of work. Decoding is the half that can be checked
        // honestly today, and it is strictly more than nothing.
        Rule::module_binary_cmd => {
            let bytes = binary_module_bytes(&pair);
            match vybe_platform_wasm::read_wasm(&bytes) {
                Ok(_) => Ok(()),
                Err(e) => Err(format!(
                    "(module binary …) failed to decode, but the fixture declares it well-formed: {e}"
                )),
            }
        }
        Rule::assert_return => {
            body.push(walk_assert_return(__w, pair)?);
            Ok(())
        }
        Rule::assert_trap | Rule::assert_instantiation_trap => {
            body.push(walk_assert_trap(__w, pair)?);
            Ok(())
        }
        // `assert_malformed` is a PARSE assertion, and parsing is the one
        // static service this front end does have: the quoted text must fail
        // to parse. Discharged here, at walk time, by running the grammar over
        // it — no runtime call, no validator.
        Rule::assert_malformed => {
            body.push(walk_assert_malformed(__w, pair)?);
            Ok(())
        }
        // `assert_invalid` asserts the module PARSES and fails VALIDATION.
        //
        // ⛔ THIS USED TO BE `Empty` — 2720 assertions that examined nothing,
        // 2053 of them inside files the scoreboard counted as PASSING. An
        // assertion discharged without looking is the same disease as a
        // checker that cannot parse its input reporting `0 problems`.
        //
        // `module_invalid_reason` implements the STRUCTURAL subset (alignment,
        // lane index, duplicate export). The 2297 "type mismatch" assertions
        // need the stack-typing pass and will FAIL here until it exists —
        // deliberately. A validator that has not been written must report as
        // absent, not as satisfied.
        Rule::assert_invalid => {
            body.push(walk_assert_invalid(__w, pair)?);
            Ok(())
        }
        // Linkability is a different property: it is settled when two modules
        // are joined, not by looking at one. Still unenforced, and still
        // counted as such — 200 assertions.
        Rule::assert_unlinkable => {
            body.push(Statement::with_span(StmtKind::Empty, to_span(&pair)));
            Ok(())
        }
        // Exhaustion/suspension are RUNTIME trap assertions with an
        // action — same lowering as assert_trap (the old message-only
        // routing dropped the action entirely).
        Rule::assert_exhaustion | Rule::assert_suspension => {
            body.push(walk_assert_trap(__w, pair)?);
            Ok(())
        }
        // `(assert_exception action)` — the action must THROW. It carries no
        // message to compare; completing normally is the failure.
        Rule::assert_exception => {
            body.push(walk_assert_exception(__w, pair)?);
            Ok(())
        }
        Rule::invoke_cmd => {
            body.push(walk_invoke_cmd(__w, pair)?);
            Ok(())
        }
        Rule::register_cmd => {
            body.push(walk_register_cmd(__w, pair)?);
            Ok(())
        }
        Rule::get_cmd => {
            body.push(walk_get_cmd(__w, pair)?);
            Ok(())
        }
        _ => Ok(()),
    }
}

// ── Component Model ───────────────────────────────────────────────────────────

/// The spec spelling of a component definition, for error messages.
///
/// The grammar rule names are prefixed (`comp_`/`core_`) to keep the two index
/// spaces apart; an error must quote what the SOURCE says, not what the parser
/// calls it.
fn component_definition_kind(rule: Rule) -> &'static str {
    match rule {
        Rule::core_module_def   => "(core module …)",
        Rule::core_instance_def => "(core instance …)",
        Rule::core_type_def     => "(core type …)",
        Rule::component         => "(component …)",
        Rule::comp_instance     => "(instance …)",
        Rule::comp_alias        => "(alias …)",
        Rule::canon             => "(canon …)",
        Rule::comp_start        => "(start …)",
        Rule::comp_import       => "(import …)",
        Rule::comp_export       => "(export …)",
        Rule::comp_value        => "(value …)",
        Rule::comp_type         => "(type …)",
        _                       => "definition",
    }
}

/// Walk `(component <id>? <definition>*)`.
///
/// The spine only. A component's core modules walk through exactly the path a
/// top-level `(module …)` takes — `walk_module` dispatches on INNER pairs and
/// never inspects its own rule, so `(core module …)` and `(module …)` are the
/// same walk — and a nested component recurses.
///
/// Every other definition kind REFUSES, naming itself. It would be far easier
/// to skip them: the grammar accepts the whole format, so a component with a
/// `canon` section parses clean and walking nothing yields a program that
/// links and runs. That is precisely the failure worth avoiding — a canon
/// section that silently evaporates leaves a `stream.read` with no element
/// type and a `thread.suspend` with no `cancellable?`, and nothing downstream
/// can tell the difference between "not declared" and "dropped on the floor".
/// A component's CORE MODULE index space.
///
/// A core module inside a component is DECLARED, not instantiated — nothing
/// runs until a `(core instance (instantiate …))` says so. That is the spec's
/// semantics and it is also what makes the `with` clause work at all: a
/// module's imports cannot be resolved until the instantiation that supplies
/// them has been read, and the instantiation comes after the module.
struct CoreModules<'i> {
    defs: Vec<Pair<'i, Rule>>,
    names: HashMap<String, u32>,
}

/// A component's COMPONENT index space.
///
/// ⛔ Exactly the `CoreModules` treatment, one level up, and for the same
/// reason: a nested `(component …)` is DECLARED, not run. Nothing inside it
/// executes until an `(instance (instantiate …))` says so, and its imports
/// cannot be resolved before the instantiation that supplies them has been
/// read — which comes after the component.
///
/// Walking a nested component inline, as this used to, runs its core modules
/// where they are written. That is wrong twice over: the modules execute
/// whether or not anything instantiates them, and a component instantiated
/// TWICE would only ever run once.
struct Components<'i> {
    defs: Vec<Pair<'i, Rule>>,
    names: HashMap<String, u32>,
}

fn walk_component(
    __w: &mut WastWalker,
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
) -> Result<HashMap<String, u32>, String> {
    // Per component, never shared with a nested one: each has its own spaces.
    let mut modules = CoreModules {
        defs: Vec::new(),
        names: HashMap::new(),
    };
    let mut components = Components {
        defs: Vec::new(),
        names: HashMap::new(),
    };
    // …and so does every OTHER index space, which is what this line used to
    // claim while only `modules` was actually scoped. The rest lived on the
    // walker and were shared with nested components, so an inner `(type …)`
    // renumbered the outer's type space and an inner alias stayed visible
    // outside it — silently, because both sides are small integers.
    //
    // `canon_section`/`canon_binder` stay SHARED on purpose: there is one
    // VM-level canon section (`VM::canon_defs`), so canonidx has to keep
    // counting across nested components. A name collision between an inner and
    // an outer binder therefore refuses, which is conservative rather than
    // wrong.
    //
    // `comp_func_space` and `comp_types` are shared for the same reason and a
    // sharper one: they are the AST PAYLOAD this walk produces, not scratch for
    // resolving names. Scoping them here restores the EMPTY outer copies on the
    // way out and discards everything the top-level component built — which is
    // what `canon lower: $callee 0 is not in the component function index space
    // (have 0)` was reporting. Only the NAME maps are per-component.
    let saved = ComponentSpaces::take(__w);
    for def in pair.into_inner() {
        match def.as_rule() {
            // The component's own `$id`. Components are not yet addressable as
            // instantiation targets, so nothing binds it.
            Rule::id => {}
            Rule::component_definition => {
                let inner = def
                    .into_inner()
                    .next()
                    .ok_or("component: empty definition")?;
                if let Err(e) = walk_component_definition(
                    __w,
                    inner,
                    body,
                    &mut modules,
                    &mut components,
                ) {
                    saved.restore(__w);
                    return Err(e);
                }
            }
            other => {
                saved.restore(__w);
                return Err(format!(
                    "component: unexpected {other:?} in a component body"
                ));
            }
        }
    }
    // This component's EXPORT TABLE, captured before the scopes unwind.
    // Instantiating a component yields an instance whose exports are exactly
    // this list (`Explainer.md:889`), so it is the return value rather than
    // something the caller digs out of the walker.
    //
    // ⛔ The funcidx values index `comp_func_space`, which is PAYLOAD and
    // therefore shared across nesting — so an inner component's export names an
    // index that is still valid out here. Scoping that space would have made
    // every returned index dangle.
    let exports = std::mem::take(&mut __w.comp_exports);
    saved.restore(__w);
    Ok(exports)
}

/// The index spaces a component owns, lifted out so a nested `(component …)`
/// gets fresh ones and the enclosing component gets its own back.
#[derive(Default)]
struct ComponentSpaces {
    core_func_space: Vec<CoreFunc>,
    core_func_index: HashMap<String, u32>,
    core_instances: Vec<CoreInstance>,
    core_instance_index: HashMap<String, u32>,
    comp_type_index: HashMap<String, u32>,
    comp_type_space: u32,
    comp_type_base: u32,
    core_type_index: HashMap<String, u32>,
    core_type_space: u32,
    comp_func_index: HashMap<String, u32>,
    comp_instances: Vec<CompInstance>,
    comp_instance_index: HashMap<String, u32>,
    comp_exports: HashMap<String, u32>,
    export_names: ExternNames,
    import_names: ExternNames,
}

impl ComponentSpaces {
    /// Move the walker's spaces out, leaving it with empty ones.
    fn take(__w: &mut WastWalker) -> Self {
        ComponentSpaces {
            core_func_space: std::mem::take(&mut __w.core_func_space),
            core_func_index: std::mem::take(&mut __w.core_func_index),
            core_instances: std::mem::take(&mut __w.core_instances),
            core_instance_index: std::mem::take(&mut __w.core_instance_index),
            comp_type_index: std::mem::take(&mut __w.comp_type_index),
            comp_type_space: std::mem::take(&mut __w.comp_type_space),
            // The nested component's own index space starts at 0 and its
            // declarations are appended, so its base is the vector's length AT
            // ENTRY. `take` leaves 0 behind, which is wrong here — the base has
            // to be SET, not cleared, which is why this is not a `take`.
            comp_type_base: std::mem::replace(
                &mut __w.comp_type_base,
                __w.comp_types.len() as u32,
            ),
            core_type_index: std::mem::take(&mut __w.core_type_index),
            core_type_space: std::mem::take(&mut __w.core_type_space),
            comp_func_index: std::mem::take(&mut __w.comp_func_index),
            comp_instances: std::mem::take(&mut __w.comp_instances),
            comp_instance_index: std::mem::take(&mut __w.comp_instance_index),
            comp_exports: std::mem::take(&mut __w.comp_exports),
            export_names: std::mem::take(&mut __w.export_names),
            import_names: std::mem::take(&mut __w.import_names),
        }
    }

    /// Put them back, discarding whatever the nested component built.
    fn restore(self, __w: &mut WastWalker) {
        __w.core_func_space = self.core_func_space;
        __w.core_func_index = self.core_func_index;
        __w.core_instances = self.core_instances;
        __w.core_instance_index = self.core_instance_index;
        __w.comp_type_index = self.comp_type_index;
        __w.comp_type_space = self.comp_type_space;
        __w.comp_type_base = self.comp_type_base;
        __w.core_type_index = self.core_type_index;
        __w.core_type_space = self.core_type_space;
        __w.comp_func_index = self.comp_func_index;
        __w.comp_instances = self.comp_instances;
        __w.comp_instance_index = self.comp_instance_index;
        __w.comp_exports = self.comp_exports;
        __w.export_names = self.export_names;
        __w.import_names = self.import_names;
    }
}

fn walk_component_definition<'i>(
    __w: &mut WastWalker,
    pair: Pair<'i, Rule>,
    body: &mut Vec<Statement>,
    modules: &mut CoreModules<'i>,
    components: &mut Components<'i>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::core_module_def => {
            // DECLARED, not walked. See `CoreModules`.
            let idx = modules.defs.len() as u32;
            if let Some(id) = pair.clone().into_inner().find(|c| c.as_rule() == Rule::id) {
                modules.names.insert(id.as_str()[1..].to_string(), idx);
            }
            modules.defs.push(pair);
            Ok(())
        }
        Rule::core_instance_def => walk_core_instance(__w, pair, body, modules),
        Rule::component => {
            // DECLARED, not walked — see `Components`. A nested component runs
            // only when an `(instance (instantiate …))` asks for it, which is
            // both the spec's ordering and what lets one be instantiated twice.
            let idx = components.defs.len() as u32;
            if let Some(id) = pair.clone().into_inner().find(|c| c.as_rule() == Rule::id) {
                components.names.insert(id.as_str()[1..].to_string(), idx);
            }
            components.defs.push(pair);
            Ok(())
        }
        Rule::comp_instance => walk_comp_instance(__w, pair, body, components),
        Rule::comp_export => walk_comp_export(__w, pair),
        Rule::comp_import => walk_comp_import(__w, pair),
        Rule::comp_alias => walk_comp_alias(__w, pair),
        Rule::comp_type => walk_comp_type(__w, pair),
        Rule::core_type_def => walk_core_type(__w, pair),
        Rule::canon => walk_canon(__w, pair),
        other => Err(format!(
            "component: {} parses but has no walk yet",
            component_definition_kind(other)
        )),
    }
}

/// `(core instance <id>? (instantiate <moduleidx> <arg>*))`, and the
/// `<core:inlineexport>*` form.
///
/// This is where a core module actually runs, where its imports are decided,
/// and where the component's CORE INSTANCE index space gets its entries. Each
/// `(with <module> (instance (export <name> (core func $b))))` binds one of
/// the module's import slots to something the COMPONENT supplies.
fn walk_core_instance<'i>(
    __w: &mut WastWalker,
    pair: Pair<'i, Rule>,
    body: &mut Vec<Statement>,
    modules: &mut CoreModules<'i>,
) -> Result<(), String> {
    let inst_name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string());
    let expr = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::core_instanceexpr)
        .ok_or("core instance: no instance expression")?;
    let Some(inst) = expr
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::core_instantiate)
    else {
        // `<core:inlineexport>*` — an instance assembled from items that
        // already exist rather than instantiated from a module. It runs no
        // code; it only names things, so all it contributes is an export
        // table.
        let mut funcs: HashMap<String, CoreFunc> = HashMap::new();
        for exp in expr.into_inner() {
            if exp.as_rule() != Rule::core_inlineexport {
                continue;
            }
            let (name, item) = core_inlineexport_item(__w, &exp)?;
            funcs.insert(name, item);
        }
        publish_core_instance(__w, inst_name, CoreInstance { funcs });
        return Ok(());
    };

    let target = inst
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)
        .ok_or("core instance: (instantiate …) names no module")?;
    let midx = resolve_idx(&target, "core module", &modules.names)?;
    let module = modules
        .defs
        .get(midx as usize)
        .cloned()
        .ok_or_else(|| {
            format!(
                "core instance: core module {midx} is not declared (have {})",
                modules.defs.len()
            )
        })?;

    // The class `walk_module` will publish this instance's exports under.
    // Computed BEFORE the walk, because `walk_module` advances `module_seq`.
    //
    // A module carrying an `$id` is published under that id, so instantiating
    // it twice republishes ONE class and the second instance shadows the
    // first. That is a real fidelity limit of the class-per-module model, not
    // something this function can paper over — an alias into either instance
    // resolves to the same class.
    let class = module
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
        .unwrap_or_else(|| {
            if __w.module_seq == 0 {
                "__wasm_module".to_string()
            } else {
                format!("__wasm_module_{}", __w.module_seq)
            }
        });

    // Build this instantiation's import wiring, then walk the module under it.
    let mut supplied: HashMap<(String, String), CoreFunc> = HashMap::new();
    for arg in inst.into_inner() {
        if arg.as_rule() != Rule::core_instantiatearg {
            continue;
        }
        let import_module = arg
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::string)
            .map(|c| unquote(c.as_str()))
            .ok_or("core instance: (with …) names no import module")?;
        let mut any = false;
        for exp in arg.clone().into_inner() {
            if exp.as_rule() != Rule::core_inlineexport {
                continue;
            }
            any = true;
            let (name, item) = core_inlineexport_item(__w, &exp)?;
            supplied.insert((import_module.clone(), name), item);
        }
        if !any {
            // `(with "m" (instance <idx>))` — every export of that instance
            // fills the slot, under its own name. Whole-instance wiring, which
            // is what a module importing several names from one namespace
            // needs.
            let iref = arg
                .clone()
                .into_inner()
                .find(|c| c.as_rule() == Rule::index)
                .ok_or("core instance: (with … (instance …)) names no instance")?;
            let iidx = resolve_idx(&iref, "core instance", &__w.core_instance_index)?;
            let src = __w.core_instances.get(iidx as usize).cloned().ok_or_else(|| {
                format!(
                    "core instance: core instance {iidx} is not declared (have {})",
                    __w.core_instances.len()
                )
            })?;
            for (ename, item) in src.funcs {
                supplied.insert((import_module.clone(), ename), item);
            }
        }
    }

    let saved = std::mem::replace(&mut __w.component_imports, supplied);
    let walked = walk_module(__w, module);
    __w.component_imports = saved;
    body.extend(walked?);

    // Publish the instance's export table. `module_exports` is keyed by class
    // and was just written by `walk_module`.
    let funcs = __w
        .module_exports
        .get(&class)
        .map(|ex| {
            ex.iter()
                .map(|(e, m)| (e.clone(), CoreFunc::Module(class.clone(), m.clone())))
                .collect()
        })
        .unwrap_or_default();
    publish_core_instance(__w, inst_name, CoreInstance { funcs });
    Ok(())
}

/// Append to the core instance index space, binding `$id` if one was written.
/// Unnamed instances still advance the space — the index is positional, and
/// skipping one would silently renumber every alias after it.
fn publish_core_instance(__w: &mut WastWalker, name: Option<String>, inst: CoreInstance) {
    let idx = __w.core_instances.len() as u32;
    __w.core_instances.push(inst);
    if let Some(n) = name {
        __w.core_instance_index.insert(n, idx);
    }
}

/// `(export <name> <core:externidx>)` → the name it publishes and the item.
fn core_inlineexport_item(
    __w: &WastWalker,
    exp: &Pair<Rule>,
) -> Result<(String, CoreFunc), String> {
    let name = exp
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::string)
        .map(|c| unquote(c.as_str()))
        .ok_or("core instance: (export …) has no name")?;
    let target = exp
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::core_externidx)
        .ok_or("core instance: (export …) names nothing")?;
    Ok((name, core_externidx_item(__w, &target)?))
}

/// `(core func $b)` / `(core func 0)` → the core function it names.
///
/// Both spellings resolve through the SAME space: `$b` through the name map,
/// a bare integer positionally. They used to be two different refusals because
/// the space did not exist — the name map was read and never written, so every
/// `$b` was "not bound" and every integer had nothing to index.
fn core_externidx_item(__w: &WastWalker, target: &Pair<Rule>) -> Result<CoreFunc, String> {
    let sort = target
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::core_sort)
        .map(|c| c.as_str().trim().to_string())
        .unwrap_or_default();
    if sort != "func" {
        return Err(format!(
            "core instance: `(core {sort} …)` — only a core FUNCTION can be supplied \
             to an import so far; the {sort} index space has no producer"
        ));
    }
    let idx = target
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)
        .ok_or("core instance: exported item names no index")?;
    let fidx = resolve_idx(&idx, "core func", &__w.core_func_index)?;
    __w.core_func_space.get(fidx as usize).cloned().ok_or_else(|| {
        format!(
            "core instance: core func {fidx} is not defined (have {})",
            __w.core_func_space.len()
        )
    })
}

/// `(alias core export <instanceidx> <name> (core <sort> <id>?))` — §5.
///
/// This is the only way a component reaches inside an instantiated core
/// module, and it is what gives the core function index space its second
/// producer (the first being a canon row's binder).
/// The names an import or export scope has already taken.
///
/// ⛔ **THE KEY IS FOLDED; THE VALUE IS THE NAME AS WRITTEN.** The exported
/// name is an identifier a host matches on, so it must survive byte-for-byte —
/// only the COLLISION KEY folds. Storing the folded form as the name would be
/// the same mistake the namespace tree makes with its lowercase-canonical keys.
///
/// Imports and exports are SEPARATE scopes: `Explainer.md:2837` says a set of
/// names "can thus all be imports (or exports) of the same component", so one
/// table each.
#[derive(Default)]
struct ExternNames {
    /// folded key → the names as the source wrote them
    ///
    /// ⛔ A **LIST**, not one name. `foo` and `[constructor]foo` fold to the
    /// same key and are legitimately both present (clause 1), so a key can hold
    /// two. With one slot the constructor OVERWRITES the label, and a second
    /// `foo` then compares only against `[constructor]foo`, is exempted by
    /// clause 1, and is ACCEPTED — while the spec lists adding `foo` twice as a
    /// validation error. A new name must be exempt against EVERY name already
    /// under its key, not just the last one written.
    by_key: HashMap<String, Vec<String>>,
    /// folded MEMBER label of an `[annotation]a.b` name → the name as written
    by_member: HashMap<String, String>,
}

/// The collision key of an extern name — `Explainer.md:2832`, clause 3:
/// "Lowercase all the acronyms … Strip any `[...]` annotation prefix".
///
/// Lowercasing the whole string is equivalent to lowercasing the acronyms: a
/// `label` is fragments that are each entirely lower-case (`word`) or entirely
/// upper-case (`acronym`), so no mixed-case fragment exists to preserve. ASCII,
/// because the `label` grammar is ASCII.
fn extern_name_key(name: &str) -> String {
    extern_name_body(name).to_ascii_lowercase()
}

/// `name` with any `[...]` annotation prefix removed.
fn extern_name_body(name: &str) -> &str {
    match name.strip_prefix('[') {
        Some(rest) => match rest.split_once(']') {
            Some((_, body)) => body,
            None => name,
        },
        None => name,
    }
}

/// The label AFTER the dot in an `[annotation]a.b` name.
///
/// ⛔ Clause 2 is written `[*]l.l` and reads as though both dotted labels must
/// equal the bare name. THE EXAMPLES SAY OTHERWISE, and they are decisive:
/// adding `bar` to a set containing `[method]foo.bar` is listed as a validation
/// error, and NOTHING ELSE in that set collides with `bar` under clause 3. So
/// the rule is that a bare `l` collides with any `[*]x.l` — the SECOND label.
///
/// The stated rationale confirms it: the point is "pathological cases where two
/// unique-in-the-component names get mapped to the same source-language
/// identifier". A method `bar` and a free function `bar` both become `bar` in a
/// generated binding; `[method]foo.bar` alongside `foo` does not.
fn extern_name_member(name: &str) -> Option<&str> {
    let body = extern_name_body(name);
    if body.len() == name.len() {
        // No annotation — a plain dotted name is not the `[*]a.b` shape.
        return None;
    }
    body.split_once('.').map(|(_, member)| member)
}

impl ExternNames {
    /// Record `name`, or refuse if it is not STRONGLY-UNIQUE against what is
    /// already here — `Explainer.md:2826`.
    fn insert(&mut self, what: &str, name: &str) -> Result<(), String> {
        let key = extern_name_key(name);
        // Clause 1: `l` and `[constructor]l` ARE strongly-unique, for the SAME
        // label — compared RAW, before folding, or `foo-bar` and
        // `[constructor]foo-BAR` would be let through when the spec's own
        // example lists that pair as an error.
        let ctor_pair =
            |a: &str, b: &str| b.strip_prefix("[constructor]").is_some_and(|l| l == a);
        for prev in self.by_key.get(&key).into_iter().flatten() {
            if !ctor_pair(prev, name) && !ctor_pair(name, prev) {
                return Err(format!(
                    "{what}: \"{name}\" is not strongly-unique against \"{prev}\" \
                     (Explainer.md:2826 — annotations stripped and acronyms lowercased, \
                     both are `{key}`)"
                ));
            }
        }
        // Clause 2, both directions.
        if let Some(member) = extern_name_member(name) {
            let mk = member.to_ascii_lowercase();
            if let Some(prev) = self.by_key.get(&mk).and_then(|v| v.first()) {
                return Err(format!(
                    "{what}: \"{name}\" is not strongly-unique against \"{prev}\" \
                     (Explainer.md:2829 — a bare `{member}` and an annotated `….{member}` \
                     generate the same binding name)"
                ));
            }
        } else if let Some(prev) = self.by_member.get(&key) {
            return Err(format!(
                "{what}: \"{name}\" is not strongly-unique against \"{prev}\" \
                 (Explainer.md:2829 — a bare `{name}` and an annotated `….{name}` \
                 generate the same binding name)"
            ));
        }
        if let Some(member) = extern_name_member(name) {
            self.by_member
                .insert(member.to_ascii_lowercase(), name.to_string());
        }
        self.by_key.entry(key).or_default().push(name.to_string());
        Ok(())
    }
}

/// `(import <externnamelit> <attribute>* bind-id(<externtype>))`.
///
/// The FOURTH producer of the component function index space, and the only one
/// that supplies no definition. `Explainer.md:2601` puts imports and exports on
/// the same footing — both "append a new element to the index space of the
/// imported/exported `sort`" — and `bind-id(<externtype>)` is where the name
/// goes: `(import "x" (func $x (type $ft)))` binds `$x` "just like Core
/// WebAssembly, as part of the `externtype`".
///
/// So the slot is created and left EMPTY, and `canon lower` naming it walks,
/// compiles and then refuses at the CALL with the linker named. That ordering
/// is the point: a component may legitimately declare an import it never calls,
/// and refusing at the declaration would reject a valid component.
fn walk_comp_import(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    let text = pair.as_str().trim();
    if let Some(a) = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::attribute)
    {
        return Err(format!(
            "import: `{}` — an attribute changes the imported name or makes a claim \
             about it, and nothing carries either yet",
            a.as_str().trim()
        ));
    }
    // The import NAME. It was not recorded at all before, so nothing checked
    // two imports for a collision — and `Explainer.md:2837` requires the same
    // STRONGLY-UNIQUE rule here as for exports, in its own scope.
    let name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::externnamelit)
        .map(|c| unquote(c.as_str()))
        .ok_or("import: no import name")?;
    __w.import_names.insert("import", &name)?;
    let ty = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::externtype)
        .ok_or("import: no external type")?;
    let tytext = ty.as_str().trim();
    // Only the `(func $x (type $ft))` shape has a producer. The rest name
    // index spaces nothing fills — refuse each by what it would need, not with
    // one blanket "unsupported".
    let head = tytext.trim_start_matches('(').trim_start();
    let sort = ty
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::comp_sort_kw)
        .map(|c| c.as_str().trim().to_string())
        .unwrap_or_default();
    if sort != "func" {
        return Err(format!(
            "import: `{text}` — {}",
            if head.starts_with("core") {
                "a `(core module …)` import needs the core module index space, which has \
                 no producer"
                    .to_string()
            } else if sort.is_empty() {
                "an inline `functype`/`componenttype`/`instancetype` has no index to \
                 record; spell it as `(func (type $ft))` naming a declared `(type …)`"
                    .to_string()
            } else {
                format!(
                    "the {sort} index space has no producer, so an imported {sort} has \
                     nowhere to be recorded"
                )
            }
        ));
    }
    let i = ty
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)
        .ok_or_else(|| {
            format!("import: `{tytext}` — a `(func …)` import must name its `(type $ft)`")
        })?;
    let ft = resolve_comp_type(__w, &i)?;
    let fidx = __w.comp_func_space.len() as u32;
    __w.comp_func_space.push(CompFunc {
        // EMPTY on purpose — see `CompFunc`. Nothing in this component defines
        // an imported function.
        canonidx: None,
        functype: Some(ft),
    });
    if let Some(id) = ty.into_inner().find(|c| c.as_rule() == Rule::id) {
        __w.comp_func_index.insert(id.as_str()[1..].to_string(), fidx);
    }
    Ok(())
}

/// `(export <id>? <externnamelit> <attribute>* <externidx> <externtype>?)`.
///
/// ⛔ **AN EXPORT APPENDS TO THE INDEX SPACE.** `Explainer.md:2601`:
///
/// > not only import definitions, but also export definitions append a new
/// > element to the index space of the imported/exported `sort` … In the case
/// > of exports, the `<id>?` right after the `export` is bound while the
/// > `<id>` inside the `<externidx>` is a reference to the preceding
/// > definition being exported.
///
/// So the two `$id`s in `(export $x "x" (func $f))` are opposites: `$f` READS
/// the function space and `$x` names the NEW entry this line adds to it.
/// Treating `$x` as a second name for `$f` would look right until something
/// indexed positionally past the export, at which point every later index is
/// off by one — the `GLOBAL_GET` shape again.
fn walk_comp_export(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    let text = pair.as_str().trim();
    // The leading `id?` is a DIRECT child; the one inside `<externidx>` is a
    // great-grandchild (externidx → index → id), so this cannot confuse them.
    let bind = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string());
    if let Some(a) = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::attribute)
    {
        // 🔗 `versionsuffix` changes the exported NAME; 🏷️ `implements` and
        // `external-id` are claims about what this export satisfies. Ignoring
        // any of them would publish a name or a guarantee the source did not
        // write.
        return Err(format!(
            "export: `{}` — an attribute changes the exported name or makes a claim \
             about it, and nothing carries either yet",
            a.as_str().trim()
        ));
    }
    if pair
        .clone()
        .into_inner()
        .any(|c| c.as_rule() == Rule::externtype)
    {
        // The trailing ascription narrows the exported item's type. Accepting
        // it without checking would report a type the export was never proven
        // to have.
        return Err(format!(
            "export: `{text}` — the trailing `<externtype>` ascribes a type to the \
             export, and nothing checks it against the item being exported"
        ));
    }
    let name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::externnamelit)
        .map(|c| unquote(c.as_str()))
        .ok_or("export: no export name")?;
    let target = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::externidx)
        .ok_or("export: names nothing")?;
    // `Explainer.md:2826` — STRONGLY-UNIQUE, not exact-string equality. The
    // table folds only the KEY; the value keeps the name as written, because
    // the exported name is what a host matches on.
    __w.export_names.insert("export", &name)?;
    let src = comp_externidx_func(__w, "export", &target)?;
    let entry = __w.comp_func_space.get(src as usize).cloned().ok_or_else(|| {
        format!("export: `{text}` names component func {src}, which is gone")
    })?;
    let fidx = __w.comp_func_space.len() as u32;
    __w.comp_func_space.push(entry);
    if let Some(n) = bind {
        __w.comp_func_index.insert(n, fidx);
    }
    __w.comp_exports.insert(name, fidx);
    Ok(())
}

/// `(alias export <instanceidx> <name> (<sort> <id>?))`.
///
/// An alias DEFINES a new index in the target space — it does not merely give
/// the aliased entity a second name — so this pushes a fresh entry. That
/// matters as soon as anything indexes positionally: `(func 1)` after an alias
/// must reach the alias, not the thing it aliases.
fn alias_component_export(
    __w: &mut WastWalker,
    pair: &Pair<Rule>,
    text: &str,
) -> Result<(), String> {
    let mut inner = pair.clone().into_inner();
    let iref = inner.next().ok_or("alias export: no instance index")?;
    let iidx = resolve_idx(&iref, "component instance", &__w.comp_instance_index)?;
    let inst = __w.comp_instances.get(iidx as usize).cloned().ok_or_else(|| {
        format!(
            "alias export: instance {iidx} is not declared (have {})",
            __w.comp_instances.len()
        )
    })?;
    let name = inner
        .next()
        .filter(|c| c.as_rule() == Rule::externnamelit)
        .map(|c| unquote(c.as_str()))
        .ok_or("alias export: no export name")?;
    let sort = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::sort_c)
        .map(|c| c.as_str().trim().to_string())
        .unwrap_or_default();
    if sort != "func" {
        return Err(format!(
            "alias export: `({sort} …)` — only a component FUNCTION can be aliased so \
             far; the {sort} index space has no producer"
        ));
    }
    let src = *inst.funcs.get(&name).ok_or_else(|| {
        let mut have: Vec<&str> = inst.funcs.keys().map(|s| s.as_str()).collect();
        have.sort_unstable();
        format!(
            "alias export: instance {iidx} exports no \"{name}\" (exports: {})",
            if have.is_empty() { "none".to_string() } else { have.join(", ") }
        )
    })?;
    let entry = __w.comp_func_space.get(src as usize).cloned().ok_or_else(|| {
        format!("alias export: `{}` names component func {src}, which is gone", text.trim())
    })?;
    let fidx = __w.comp_func_space.len() as u32;
    __w.comp_func_space.push(entry);
    if let Some(id) = pair.clone().into_inner().find(|c| c.as_rule() == Rule::id) {
        __w.comp_func_index.insert(id.as_str()[1..].to_string(), fidx);
    }
    Ok(())
}

/// `(instance <id>? <instanceexpr>)` — the COMPONENT instance space.
///
/// `<instanceexpr>` is a rule of its own, not a silent one, so the
/// `<inlineexport>*` pairs are GRANDCHILDREN. Reading them as direct children
/// finds none and publishes an empty instance, and the alias that follows then
/// reports `exports no "…"` — a walk bug wearing a source error's clothes.
/// `walk_core_instance` descends through `core_instanceexpr` for the same
/// reason; this follows it.
/// `(instance <id>? (instantiate <componentidx> <arg>*))`.
///
/// ⛔ **THIS IS WHERE A NESTED COMPONENT RUNS.** It was declared, not walked
/// (`Components`), so nothing inside it has executed yet — its core modules,
/// its canon rows and its instantiations all happen here, when something asks
/// for an instance. That ordering is the spec's, and it is also what makes a
/// component instantiable TWICE: walking it inline ran it once, at its
/// declaration, whether or not anyone wanted it.
///
/// The result is an instance whose exports are the component's export list
/// (`Explainer.md:889` — a component type carries two named lists, imports and
/// exports), which is exactly what `walk_component` returns.
fn instantiate_component<'i>(
    __w: &mut WastWalker,
    inst: &Pair<'i, Rule>,
    name: Option<String>,
    body: &mut Vec<Statement>,
    components: &mut Components<'i>,
) -> Result<(), String> {
    let target = inst
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)
        .ok_or("instance: (instantiate …) names no component")?;
    let cidx = resolve_idx(&target, "component", &components.names)?;
    let def = components.defs.get(cidx as usize).cloned().ok_or_else(|| {
        format!(
            "instance: component {cidx} is not declared (have {})",
            components.defs.len()
        )
    })?;
    // ⛔ A `(with …)` supplies one of the component's IMPORTS, and an imported
    // component function has no defining row anywhere — the one case that
    // genuinely needs the component linker. Accepting the clause and ignoring
    // it would instantiate with the import unsupplied and then fail at the
    // CALL, naming the wrong thing.
    if let Some(arg) = inst
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::instantiatearg)
    {
        return Err(format!(
            "instance: `{}` supplies an IMPORT to a component instantiation, which needs \
             the component linker (see cmplan.md §Deferred to export). A component with \
             no imports instantiates without any `(with …)`",
            arg.as_str().trim()
        ));
    }
    // Run it. `walk_component` installs its own scopes, so the inner
    // component's names cannot leak out here and the outer's are invisible
    // inside.
    let exports = walk_component(__w, def, body)?;
    let idx = __w.comp_instances.len() as u32;
    __w.comp_instances.push(CompInstance { funcs: exports });
    if let Some(n) = name {
        __w.comp_instance_index.insert(n, idx);
    }
    Ok(())
}

fn walk_comp_instance<'i>(
    __w: &mut WastWalker,
    pair: Pair<'i, Rule>,
    body: &mut Vec<Statement>,
    components: &mut Components<'i>,
) -> Result<(), String> {
    let name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string());
    let expr = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::instanceexpr)
        .ok_or("instance: no instance expression")?;
    if let Some(inst) = expr
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::comp_instantiate)
    {
        return instantiate_component(__w, &inst, name, body, components);
    }

    // `<inlineexport>*` — an instance assembled from items that already exist.
    // It runs no code; all it contributes is an export table.
    let mut funcs: HashMap<String, u32> = HashMap::new();
    for exp in expr.into_inner() {
        if exp.as_rule() != Rule::inlineexport {
            continue;
        }
        if exp
            .clone()
            .into_inner()
            .any(|c| c.as_rule() == Rule::versionsuffix)
        {
            // 🔗 The suffix is part of the exported NAME, so dropping it would
            // make the alias below match on a name the component never
            // exported.
            return Err(format!(
                "instance: `{}` — a `(versionsuffix …)` changes the exported name and \
                 nothing carries it yet",
                exp.as_str().trim()
            ));
        }
        let ename = exp
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::externnamelit)
            .map(|c| unquote(c.as_str()))
            .ok_or("instance: (export …) has no name")?;
        let target = exp
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::externidx)
            .ok_or("instance: (export …) names nothing")?;
        funcs.insert(ename, comp_externidx_func(__w, "instance", &target)?);
    }
    let idx = __w.comp_instances.len() as u32;
    __w.comp_instances.push(CompInstance { funcs });
    if let Some(n) = name {
        __w.comp_instance_index.insert(n, idx);
    }
    Ok(())
}

/// Resolve an `<externidx>` to a COMPONENT function index.
///
/// The grammar's three alternatives are `(core <sort> …)`, `(<sort> …)` and a
/// bare index. Only the middle one naming `func` can be honoured: a component
/// instance cannot export a core item, and a bare index does not say which
/// index space it means — resolving it as a function would be a guess that
/// looks right whenever the two spaces happen to line up.
fn comp_externidx_func(
    __w: &WastWalker,
    what: &str,
    target: &Pair<Rule>,
) -> Result<u32, String> {
    let text = target.as_str().trim();
    if text.trim_start_matches('(').trim_start().starts_with("core") {
        return Err(format!(
            "{what}: `{text}` — a component names COMPONENT items here; a core item has \
             to be lifted first"
        ));
    }
    let Some(sort) = target
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::comp_sort_kw)
    else {
        return Err(format!(
            "{what}: `{text}` — the item must name its sort, as `(func {text})`; a bare \
             index does not say which index space it indexes"
        ));
    };
    let sort = sort.as_str().trim();
    if sort != "func" {
        return Err(format!(
            "{what}: `({sort} …)` — only a component FUNCTION is reachable so far; the \
             {sort} index space has no producer"
        ));
    }
    if target.clone().into_inner().any(|c| c.as_rule() == Rule::string) {
        return Err(format!(
            "{what}: `{text}` — the trailing name form addresses an export PATH through \
             an instance, which needs the same resolution `(alias export …)` does; spell \
             it as an alias instead"
        ));
    }
    let idx = target
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)
        .ok_or_else(|| format!("{what}: the item names no index"))?;
    let fidx = resolve_idx(&idx, "component func", &__w.comp_func_index)?;
    if fidx as usize >= __w.comp_func_space.len() {
        return Err(format!(
            "{what}: component func {fidx} is not defined (have {})",
            __w.comp_func_space.len()
        ));
    }
    Ok(fidx)
}

fn walk_comp_alias(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    let text = pair.as_str();
    let mut inner = pair.clone().into_inner();
    let Some(iref) = inner.next() else {
        return Err("alias: empty".to_string());
    };
    // Only the `core export` form has `(index, string)` leading its children;
    // `alias export` leads with an externnamelit and `alias outer` with two
    // indices. Discriminate on the source text, which is what distinguishes
    // them in the grammar's three alternatives.
    let trimmed = text.trim_start_matches('(').trim_start();
    if !trimmed.starts_with("alias") {
        return Err(format!("alias: unexpected `{}`", text.trim()));
    }
    let rest = trimmed["alias".len()..].trim_start();
    if rest.starts_with("export") {
        return alias_component_export(__w, &pair, text);
    }
    if !rest.starts_with("core") {
        return Err(format!(
            "component: `{}` — `alias outer` needs enclosing-component scopes, which \
             no space carries yet",
            text.trim()
        ));
    }
    let iidx = resolve_idx(&iref, "core instance", &__w.core_instance_index)?;
    let inst = __w.core_instances.get(iidx as usize).cloned().ok_or_else(|| {
        format!(
            "alias core export: core instance {iidx} is not declared (have {})",
            __w.core_instances.len()
        )
    })?;
    let name = inner
        .next()
        .filter(|c| c.as_rule() == Rule::string)
        .map(|c| unquote(c.as_str()))
        .ok_or("alias core export: no export name")?;
    let item = inst.funcs.get(&name).cloned().ok_or_else(|| {
        let mut have: Vec<&str> = inst.funcs.keys().map(|s| s.as_str()).collect();
        have.sort_unstable();
        format!(
            "alias core export: core instance {iidx} exports no \"{name}\" (exports: {})",
            if have.is_empty() { "none".to_string() } else { have.join(", ") }
        )
    })?;
    let sort = inner
        .clone()
        .find(|c| c.as_rule() == Rule::core_sort)
        .map(|c| c.as_str().trim().to_string())
        .unwrap_or_default();
    if sort != "func" {
        return Err(format!(
            "alias core export: `(core {sort} …)` — only a core FUNCTION can be \
             aliased so far; the {sort} index space has no producer"
        ));
    }
    let fidx = __w.core_func_space.len() as u32;
    __w.core_func_space.push(item);
    if let Some(id) = inner.find(|c| c.as_rule() == Rule::id) {
        __w.core_func_index.insert(id.as_str()[1..].to_string(), fidx);
    }
    Ok(())
}


/// `(type <id>? <deftype>)` — one entry in the component's TYPE space.
///
/// The body is not walked. What a later `canon` row needs from a type
/// declaration is its INDEX, and the index is positional: every `(type …)`
/// advances the space whether or not it is named. Skipping an unnamed one
/// would silently renumber every row after it.
fn walk_comp_type(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    // LOCAL for the source's own numbering, GLOBAL for the payload vector the
    // VM indexes. The name map stores the GLOBAL one so nothing has to
    // remember to add the base at every lookup.
    let idx = __w.comp_type_base + __w.comp_type_space;
    __w.comp_type_space += 1;
    if let Some(id) = pair.clone().into_inner().find(|c| c.as_rule() == Rule::id) {
        __w.comp_type_index.insert(id.as_str()[1..].to_string(), idx);
    }
    // The body IS walked now. It used to advance a bare counter and drop the
    // declaration, which is why `VM::canon_functypes` was empty and
    // `canon lift` trapped on `$ft` even when the source declared the type.
    // The binder is what NAMES a resource. `component::ValType::Own` holds a
    // STRING and the source writes an INDEX, so this is the only place the two
    // can be connected — `(type $file (resource (rep i32)))` makes
    // `(own $file)` mean the resource named `file`. Read BEFORE `into_inner()`
    // consumes the pair.
    let bound = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string());
    let body = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::deftype)
        .and_then(|d| d.into_inner().next());
    let decl = match body {
        Some(b) => match type_decl(__w, &b)? {
            vybe_ast::canon::TypeDecl::Resource(_) => {
                vybe_ast::canon::TypeDecl::Resource(bound)
            }
            other => other,
        },
        None => vybe_ast::canon::TypeDecl::Opaque("type".to_string()),
    };
    __w.comp_types.push(decl);
    debug_assert_eq!(
        __w.comp_types.len() as u32,
        __w.comp_type_base + __w.comp_type_space
    );
    Ok(())
}

/// One `<deftype>` alternative → the AST's record of it.
fn type_decl(
    __w: &WastWalker,
    b: &Pair<Rule>,
) -> Result<vybe_ast::canon::TypeDecl, String> {
    use vybe_ast::canon::TypeDecl;
    Ok(match b.as_rule() {
        Rule::functype_c => {
            let mut params = Vec::new();
            let mut result = None;
            for c in b.clone().into_inner() {
                match c.as_rule() {
                    Rule::func_param_c => {
                        let mut it = c.into_inner();
                        let name = it
                            .next()
                            .filter(|x| x.as_rule() == Rule::labellit)
                            .map(|x| unquote(x.as_str()))
                            .ok_or("component functype: (param …) has no label")?;
                        let ty = it
                            .find(|x| x.as_rule() == Rule::valtype)
                            .ok_or("component functype: (param …) has no type")?;
                        params.push((name, val_spec(__w, &ty)?));
                    }
                    Rule::func_result_c => {
                        let ty = c
                            .into_inner()
                            .find(|x| x.as_rule() == Rule::valtype)
                            .ok_or("component functype: (result …) has no type")?;
                        result = Some(val_spec(__w, &ty)?);
                    }
                    _ => {}
                }
            }
            TypeDecl::Func { params, result }
        }
        // `resourcetype`, `componenttype`, `instancetype` occupy their index and
        // nothing reads their shape yet. Recorded rather than skipped: skipping
        // one renumbers every later typeidx.
        // The NAME is filled in by `walk_comp_type`, which is where the
        // binder is in scope.
        Rule::resourcetype => TypeDecl::Resource(None),
        Rule::componenttype => TypeDecl::Opaque("component".to_string()),
        Rule::instancetype => TypeDecl::Opaque("instance".to_string()),
        Rule::defvaltype => TypeDecl::Value(val_spec_inner(__w, b)?),
        other => TypeDecl::Opaque(format!("{other:?}")),
    })
}

/// Resolve a `<typeidx>` in the COMPONENT type space.
///
/// ⛔ A bare integer is the component's OWN index — the spec gives every
/// component a type index space starting at 0 — while `comp_types` is one flat
/// vector for the whole program because `VM::canon_types` is one table. So a
/// positional reference has to be rebased and a `$name` must not be: the name
/// map already stores the global index.
///
/// One function so the base is applied in exactly one place. Adding it at each
/// call site is how `case_sensitive` went wrong — 33 sites, 23 forgot.
fn resolve_comp_type(__w: &WastWalker, pair: &Pair<Rule>) -> Result<u32, String> {
    let raw = pair.as_str().trim();
    if raw.starts_with('$') {
        return resolve_idx(pair, "component type", &__w.comp_type_index);
    }
    let local = resolve_idx(pair, "component type", &__w.comp_type_index)?;
    Ok(__w.comp_type_base + local)
}

/// `<valtype>` → the AST's spelling of it. `valtype ::= <typeidx> | <defvaltype>`.
fn val_spec(__w: &WastWalker, v: &Pair<Rule>) -> Result<vybe_ast::canon::ValSpec, String> {
    use vybe_ast::canon::ValSpec;
    let inner = v
        .clone()
        .into_inner()
        .next()
        .ok_or("component type: empty valtype")?;
    match inner.as_rule() {
        Rule::defvaltype => val_spec_inner(__w, &inner),
        // A typeidx used as a value type.
        Rule::index => Ok(ValSpec::Ref(resolve_comp_type(__w, &inner)?)),
        other => Err(format!("component type: unexpected {other:?} in a valtype")),
    }
}

/// The `<defvaltype>` alternatives.
fn val_spec_inner(
    __w: &WastWalker,
    d: &Pair<Rule>,
) -> Result<vybe_ast::canon::ValSpec, String> {
    use vybe_ast::canon::ValSpec;
    let text = d.as_str().trim();
    let mut children = d.clone().into_inner().peekable();
    // A primitive is an ATOMIC rule with no children — the whole alternative is
    // the token. Every composite opens with `(`.
    if !text.starts_with('(') {
        return Ok(ValSpec::Prim(text.to_string()));
    }
    // Which composite it is, is not in the pair's rule (they all share
    // `defvaltype`), so it comes from the keyword after the paren — the same
    // discrimination the grammar's alternatives make.
    let head: String = text[1..]
        .trim_start()
        .chars()
        .take_while(|c| c.is_ascii_alphanumeric() || *c == '-')
        .collect();
    let vals = |w: &WastWalker, it: &mut dyn Iterator<Item = Pair<Rule>>| -> Result<Vec<ValSpec>, String> {
        let mut out = Vec::new();
        for c in it {
            if c.as_rule() == Rule::valtype {
                out.push(val_spec(w, &c)?);
            }
        }
        Ok(out)
    };
    Ok(match head.as_str() {
        "record" => {
            let mut fields = Vec::new();
            for c in children {
                if c.as_rule() != Rule::record_field {
                    continue;
                }
                let mut it = c.into_inner();
                let name = it
                    .next()
                    .filter(|x| x.as_rule() == Rule::labellit)
                    .map(|x| unquote(x.as_str()))
                    .ok_or("record: (field …) has no label")?;
                let ty = it
                    .find(|x| x.as_rule() == Rule::valtype)
                    .ok_or("record: (field …) has no type")?;
                fields.push((name, val_spec(__w, &ty)?));
            }
            ValSpec::Record(fields)
        }
        "variant" => {
            let mut cases = Vec::new();
            for c in children {
                if c.as_rule() != Rule::variant_case {
                    continue;
                }
                let mut it = c.into_inner();
                let name = it
                    .next()
                    .filter(|x| x.as_rule() == Rule::labellit)
                    .map(|x| unquote(x.as_str()))
                    .ok_or("variant: (case …) has no label")?;
                let payload = match it.find(|x| x.as_rule() == Rule::valtype) {
                    Some(t) => Some(val_spec(__w, &t)?),
                    None => None,
                };
                cases.push((name, payload));
            }
            ValSpec::Variant(cases)
        }
        "list" => {
            let mut it = d.clone().into_inner();
            let elem = it
                .find(|x| x.as_rule() == Rule::valtype)
                .ok_or("list: no element type")?;
            let elem = Box::new(val_spec(__w, &elem)?);
            // 🔧 the fixed-length form carries a trailing count.
            match it.find(|x| x.as_rule() == Rule::integer) {
                Some(n) => ValSpec::ListFixed(
                    elem,
                    n.as_str()
                        .replace('_', "")
                        .parse::<u32>()
                        .map_err(|_| format!("list: `{}` is not a length", n.as_str()))?,
                ),
                None => ValSpec::List(elem),
            }
        }
        "tuple" => ValSpec::Tuple(vals(__w, &mut children)?),
        "option" => ValSpec::Option(Box::new(
            vals(__w, &mut children)?
                .into_iter()
                .next()
                .ok_or("option: no element type")?,
        )),
        "result" => {
            // `(result <valtype>? (error <valtype>)?)` — each side independent.
            let mut ok = None;
            let mut err = None;
            for c in children {
                match c.as_rule() {
                    Rule::valtype => ok = Some(Box::new(val_spec(__w, &c)?)),
                    Rule::result_err => {
                        let t = c
                            .into_inner()
                            .find(|x| x.as_rule() == Rule::valtype)
                            .ok_or("result: (error …) has no type")?;
                        err = Some(Box::new(val_spec(__w, &t)?));
                    }
                    _ => {}
                }
            }
            ValSpec::Result(ok, err)
        }
        // 🗺️ `(map <keytype> <valtype>)`. The KEY is its own atomic rule, not a
        // `valtype` — the spec restricts a map key to the primitives that have
        // a total ordering, so no float and no `error-context`.
        //
        // ⛔ That is why this cannot go through `vals`, which collects
        // `Rule::valtype` children and nothing else. It used to, and so it
        // skipped the key entirely: `it.next()` returned the VALUE type and
        // called it the key, then found nothing left and refused with
        // "map: no value type" — a message that pointed at the half of the
        // source that was present.
        "map" => {
            let mut k = None;
            let mut v = None;
            for c in children {
                match c.as_rule() {
                    // Atomic, so the whole primitive name is the text.
                    Rule::keytype => k = Some(ValSpec::Prim(c.as_str().trim().to_string())),
                    Rule::valtype => v = Some(val_spec(__w, &c)?),
                    _ => {}
                }
            }
            ValSpec::Map(
                Box::new(k.ok_or("map: no key type")?),
                Box::new(v.ok_or("map: no value type")?),
            )
        }
        "stream" | "future" => {
            let elem = vals(__w, &mut children)?.into_iter().next().map(Box::new);
            if head == "stream" {
                ValSpec::Stream(elem)
            } else {
                ValSpec::Future(elem)
            }
        }
        "flags" | "enum" => {
            let labels: Vec<String> = children
                .filter(|c| c.as_rule() == Rule::labellit)
                .map(|c| unquote(c.as_str()))
                .collect();
            if head == "flags" {
                ValSpec::Flags(labels)
            } else {
                ValSpec::Enum(labels)
            }
        }
        "own" | "borrow" => {
            let i = children
                .find(|c| c.as_rule() == Rule::index)
                .ok_or_else(|| format!("{head}: no resource type index"))?;
            let n = resolve_comp_type(__w, &i)?;
            if head == "own" {
                ValSpec::Own(n)
            } else {
                ValSpec::Borrow(n)
            }
        }
        // `keytype` is an atomic primitive inside `(map …)`; it reaches here
        // only if the grammar gains an alternative this match has not learned.
        other => {
            let _ = children.peek();
            return Err(format!(
                "component type: `{other}` is not a defvaltype this walker knows"
            ));
        }
    })
}

/// `(core type <id>? …)` — the CORE type space, which `thread.new-indirect`
/// and the spawn rows index. A different space from `walk_comp_type`'s.
fn walk_core_type(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    let idx = __w.core_type_space;
    __w.core_type_space += 1;
    if let Some(id) = pair.into_inner().find(|c| c.as_rule() == Rule::id) {
        __w.core_type_index.insert(id.as_str()[1..].to_string(), idx);
    }
    Ok(())
}

// ── The canon section ────────────────────────────────────────────────────────

/// Resolve an `index` against a named index space.
///
/// A bare integer is the index. A `$id` must already be bound — an unbound one
/// REFUSES rather than defaulting to 0. Substituting a plausible index is the
/// `GLOBAL_GET` defect exactly: it links, it runs, and it addresses the wrong
/// entity with nothing to detect it.
fn resolve_idx(
    pair: &Pair<Rule>,
    space_name: &str,
    names: &HashMap<String, u32>,
) -> Result<u32, String> {
    let raw = pair.as_str().trim();
    match raw.strip_prefix('$') {
        Some(name) => names.get(name).copied().ok_or_else(|| {
            format!("canon: `${name}` is not bound in the {space_name} index space")
        }),
        None => raw
            .replace('_', "")
            .parse::<u32>()
            .map_err(|_| format!("canon: `{raw}` is not a {space_name} index")),
    }
}

/// The `index` inside a `(core func $i "name")`-style reference, or a bare one.
fn idx_child<'a>(pair: &'a Pair<Rule>) -> Option<Pair<'a, Rule>> {
    if pair.as_rule() == Rule::index {
        return Some(pair.clone());
    }
    pair.clone().into_inner().find(|c| c.as_rule() == Rule::index)
}

/// Is this rule one of the atomic keyword rules that names a canon row?
fn is_canon_op(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::canon_nullary_op
            | Rule::canon_ty_op
            | Rule::canon_ty_opts_op
            | Rule::canon_ty_async_op
            | Rule::canon_ctx_op
            | Rule::canon_task_return_op
            | Rule::canon_ws_mem_op
            | Rule::canon_async_only_op
            | Rule::canon_cancellable_op
            | Rule::canon_errctx_op
            | Rule::canon_thread_new_indirect_op
            | Rule::canon_spawn_ref_op
            | Rule::canon_spawn_indirect_op
            | Rule::canon_avail_par_op
    )
}

/// `(core func <id>?)` — the binder a canon row publishes.
fn canon_binder_name(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
}

/// One `<canonopt>` onto the row's options.
fn apply_canonopt(
    __w: &WastWalker,
    opts: &mut vybe_ast::canon::CanonOptions,
    pair: Pair<Rule>,
) -> Result<(), String> {
    let text = pair.as_str();
    // Each option resolves against ITS OWN space. `(memory m)` is a core
    // memidx and `(realloc f)` a core funcidx; consulting one map for both is
    // the one-integer-two-index-spaces defect, and it fails silently — the
    // name resolves, to the wrong entity.
    let idx = |p: &Pair<Rule>,
               space: &str,
               names: &HashMap<String, u32>|
     -> Result<u32, String> {
        idx_child(p)
            .ok_or_else(|| format!("canon: {space} option carries no index"))
            .and_then(|i| resolve_idx(&i, space, names))
    };
    for child in pair.clone().into_inner() {
        match child.as_rule() {
            Rule::string_encoding_opt => {
                let enc = child.as_str().trim_start_matches("string-encoding=");
                opts.string_encoding = Some(enc.to_string());
                return Ok(());
            }
            Rule::async_kw => {
                opts.is_async = true;
                return Ok(());
            }
            Rule::core_memidx_ref => {
                opts.memory = Some(idx(&child, "core memory", &__w.core_memory_index)?);
                return Ok(());
            }
            Rule::core_funcidx_ref => {
                // Which option this is is decided by the KEYWORD, not the
                // operand shape — `realloc`, `post-return` and `callback` all
                // take a core funcidx.
                let v = Some(idx(&child, "core func", &__w.core_func_index)?);
                if text.starts_with("(realloc") {
                    opts.realloc = v;
                } else if text.starts_with("(post-return") {
                    opts.post_return = v;
                } else if text.starts_with("(callback") {
                    opts.callback = v;
                } else {
                    return Err(format!("canon: unrecognised option `{text}`"));
                }
                return Ok(());
            }
            _ => {}
        }
    }
    Err(format!("canon: unrecognised option `{text}`"))
}

/// `(canon …)` — one row of the canon section.
fn walk_canon(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<(), String> {
    let body = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::canon_body)
        .ok_or("canon: empty definition")?;
    let row = body.into_inner().next().ok_or("canon: empty body")?;
    let decl = match row.as_rule() {
        Rule::canon_lift => walk_canon_lift(__w, row)?,
        Rule::canon_lower => walk_canon_lower(__w, row)?,
        Rule::canon_builtin => walk_canon_builtin(__w, row)?,
        other => return Err(format!("canon: unexpected {other:?}")),
    };
    // A row's binder is how a core module reaches it. Recorded BEFORE the push
    // so the canonidx is this row's own index.
    if let Some(binder) = decl.binder.clone() {
        let canonidx = __w.canon_section.len() as u32;
        if let Some((prev, at)) = __w
            .canon_binder
            .insert(binder.clone(), (decl.builtin.clone(), canonidx))
        {
            return Err(format!(
                "canon: `${binder}` is already bound by `{prev}` at canonidx {at}; \
                 a core function binder names one definition"
            ));
        }
        // A canon definition DEFINES a core function — `canon lower` and every
        // canonical built-in do. Recording it in the core function index space
        // is what makes `(core func $b)` and a positional `(core func 0)` name
        // the same entity, and what a later `(alias core export …)` entry sits
        // beside. Until this existed the space was read in six places and
        // written in none, so every `$id` in it was "not bound" no matter what
        // the source declared.
        let fidx = __w.core_func_space.len() as u32;
        __w.core_func_space
            .push(CoreFunc::Canon(decl.builtin.clone(), canonidx));
        __w.core_func_index.insert(binder, fidx);
    }
    __w.canon_section.push(decl);
    Ok(())
}

fn walk_canon_lift(
    __w: &mut WastWalker,
    row: Pair<Rule>,
) -> Result<vybe_ast::canon::CanonDecl, String> {
    use vybe_ast::canon::{CanonCallee, CanonDecl};
    let mut decl = CanonDecl::new("lift", None);
    let mut bind: Option<String> = None;
    for child in row.into_inner() {
        match child.as_rule() {
            Rule::core_funcidx_ref => {
                let i = idx_child(&child).ok_or("canon lift: $callee carries no index")?;
                let raw = i.as_str().trim().to_string();
                // A NAMED callee resolves through the component's core function
                // index space. The canon section records `$callee` as a plain
                // number, and `call_canon_callee` reads that number as a CHUNK
                // index — a different space, assigned by the compiler. Writing
                // this space's position there would call whichever chunk happens
                // to sit at it: two small integers meaning different things
                // depending on who reads them, which is the exact defect the
                // canon section exists to remove.
                //
                // So it refuses, and names the dependency rather than the
                // source. A BARE integer is left alone — under the VM's current
                // contract that number IS a chunk index, and every canon test
                // that works today writes one.
                if raw.starts_with('$') {
                    let fidx = resolve_idx(&i, "core func", &__w.core_func_index)?;
                    let item = __w.core_func_space.get(fidx as usize).ok_or_else(|| {
                        format!(
                            "canon lift: core func {fidx} is not defined (have {})",
                            __w.core_func_space.len()
                        )
                    })?;
                    decl.callee = Some(match item {
                        // Hand over the NAMES. The number the runtime wants is a
                        // chunk index only the compiler assigns, so writing this
                        // space's position here would call whichever chunk
                        // happened to sit at it.
                        CoreFunc::Module(class, method) => CanonCallee::CoreExport {
                            class: class.clone(),
                            method: method.clone(),
                        },
                        // Lifting a canonical definition's core function is
                        // spec-legal (a `canon lower` result can be lifted
                        // again), but a canon row compiles to no chunk, so
                        // there is nothing for `$callee` to name.
                        CoreFunc::Canon(builtin, canonidx) => {
                            return Err(format!(
                                "canon lift: `{raw}` names the core function defined by \
                                 `canon {builtin}` at canonidx {canonidx}. A canonical \
                                 definition compiles to no chunk, so there is nothing \
                                 for `$callee` to name"
                            ))
                        }
                    });
                    continue;
                }
                decl.callee = Some(CanonCallee::Core(resolve_idx(
                    &i,
                    "core func",
                    &__w.core_func_index,
                )?));
            }
            Rule::canonopt => apply_canonopt(__w, &mut decl.opts, child)?,
            Rule::externtype => {
                // `bind-id(<externtype>)` — `(func $lifted (type $ft))`. The
                // binder names the COMPONENT function this row defines, which
                // is a different space from the `(core func $b)` binder every
                // other row carries.
                bind = child
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
                // `(func (type $ft))` names the lifted signature by INDEX,
                // which is what the canon section records. The inline
                // `(func (param …) (result …))` form states the type in place
                // and has no index to record, so it refuses rather than
                // recording `None` and letting a downstream `require_type`
                // report a missing immediate the source did in fact supply.
                let named = child
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::index);
                match named {
                    Some(i) => {
                        decl.functype =
                            Some(resolve_comp_type(__w, &i)?)
                    }
                    None => {
                        return Err(format!(
                            "canon lift: the lifted type is written inline (`{}`); \
                             name it with `(type $t …)` and lift `(func (type $t))` \
                             so the canon section can record its index",
                            child.as_str().trim()
                        ))
                    }
                }
            }
            _ => {}
        }
    }
    // A `canon lift` DEFINES a component function whether or not it is named,
    // so the space advances either way — skipping the unnamed ones would
    // silently renumber every `canon lower` after them. The canonidx is this
    // row's own: `walk_canon` pushes it immediately after we return.
    let fidx = __w.comp_func_space.len() as u32;
    __w.comp_func_space.push(CompFunc {
        canonidx: Some(__w.canon_section.len() as u32),
        functype: decl.functype,
    });
    if let Some(name) = bind {
        if let Some(prev) = __w.comp_func_index.insert(name.clone(), fidx) {
            return Err(format!(
                "canon lift: `${name}` is already bound to component func {prev}; \
                 a binder names one function"
            ));
        }
    }
    Ok(decl)
}

fn walk_canon_lower(
    __w: &mut WastWalker,
    row: Pair<Rule>,
) -> Result<vybe_ast::canon::CanonDecl, String> {
    use vybe_ast::canon::{CanonCallee, CanonDecl};
    let mut decl = CanonDecl::new("lower", None);
    for child in row.into_inner() {
        match child.as_rule() {
            Rule::funcidx_ref => {
                let i = idx_child(&child).ok_or("canon lower: $callee carries no index")?;
                // A COMPONENT funcidx — a different space from `lift`'s.
                let fidx = resolve_idx(&i, "component func", &__w.comp_func_index)?;
                let f = __w.comp_func_space.get(fidx as usize).ok_or_else(|| {
                    format!(
                        "canon lower: component func {fidx} is not in the component \
                         function index space (have {}); `canon lift`, \
                         `(alias export …)`, `(export …)` and `(import …)` each add \
                         one — an import adds the SLOT without filling it",
                        __w.comp_func_space.len()
                    )
                })?;
                // ⛔ `ft` comes FROM THE CALLEE. `Binary.md:297` is
                // `0x01 0x00 f:<funcidx> opts:<opts>` — the lower row carries
                // no `ft` immediate at all, because `canon_lower(callee, ft,
                // opts, flat_args)` receives it from the function being
                // lowered. Re-deriving it here instead of taking the lift
                // row's own would let the two disagree.
                decl.functype = f.functype;
                decl.callee = Some(CanonCallee::Component(fidx));
            }
            Rule::canonopt => apply_canonopt(__w, &mut decl.opts, child)?,
            Rule::core_func_bind => decl.binder = canon_binder_name(&child),
            _ => {}
        }
    }
    Ok(decl)
}

fn walk_canon_builtin(
    __w: &mut WastWalker,
    row: Pair<Rule>,
) -> Result<vybe_ast::canon::CanonDecl, String> {
    use vybe_ast::canon::CanonDecl;
    let shape = row.into_inner().next().ok_or("canon: empty builtin")?;
    let kind = shape.as_rule();
    let mut decl = CanonDecl::default();
    // Index immediates in SOURCE ORDER. Which space each belongs to is decided
    // by the row, below — that is the whole point of spelling the rows out by
    // shape instead of accepting `<name> <arg>*`.
    let mut idxs: Vec<u32> = Vec::new();
    let mut ctx_valtype: Option<String> = None;
    let mut ctx_slot: Option<u32> = None;

    for child in shape.into_inner() {
        let r = child.as_rule();
        if is_canon_op(r) {
            decl.builtin = child.as_str().trim().to_string();
            continue;
        }
        match r {
            Rule::cancellable_kw => decl.cancellable = true,
            Rule::shared_kw => decl.shared = true,
            // A row-level `async?`. An `async` inside `<canonopt>` is nested in
            // a `canonopt` pair and never reaches here.
            Rule::async_kw => decl.is_async = true,
            Rule::canonopt => apply_canonopt(__w, &mut decl.opts, child)?,
            Rule::core_func_bind => decl.binder = canon_binder_name(&child),
            Rule::core_memidx_ref => {
                let i = idx_child(&child).ok_or("canon: (memory …) carries no index")?;
                decl.opts.memory = Some(resolve_idx(&i, "core memory", &__w.core_memory_index)?);
            }
            Rule::index => {
                idxs.push(resolve_comp_type(__w, &child)?)
            }
            Rule::core_typeidx_ref => {
                let i = idx_child(&child).ok_or("canon: core type ref carries no index")?;
                idxs.push(resolve_idx(&i, "core type", &__w.core_type_index)?);
            }
            Rule::core_tableidx_ref => {
                let i = idx_child(&child).ok_or("canon: core table ref carries no index")?;
                idxs.push(resolve_idx(&i, "core table", &__w.core_table_index)?);
            }
            Rule::valtype => ctx_valtype = Some(child.as_str().trim().to_string()),
            Rule::integer => {
                ctx_slot = Some(
                    child
                        .as_str()
                        .replace('_', "")
                        .parse::<u32>()
                        .map_err(|_| format!("canon: `{}` is not a u32", child.as_str()))?,
                )
            }
            Rule::func_result_c => {
                // `task.return (result <valtype>)?` — the result list. Only an
                // indexed valtype has a number to record here.
                if let Some(i) = child.clone().into_inner().find(|c| c.as_rule() == Rule::index) {
                    decl.results
                        .push(resolve_comp_type(__w, &i)?);
                }
            }
            _ => {}
        }
    }

    // Assign the positional immediates to their fields, per row.
    match kind {
        Rule::canon_ty | Rule::canon_ty_opts | Rule::canon_ty_async | Rule::canon_spawn_ref => {
            decl.ty = idxs.first().copied();
        }
        Rule::canon_thread_new_indirect | Rule::canon_spawn_indirect => {
            decl.ty = idxs.first().copied();
            decl.table = idxs.get(1).copied();
        }
        Rule::canon_ctx => {
            let v = ctx_valtype.ok_or("canon context.*: no valtype immediate")?;
            let i = ctx_slot.ok_or("canon context.*: no slot immediate")?;
            decl.context = Some((v, i));
        }
        _ => {}
    }
    Ok(decl)
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
fn find_first_id(pair: &Pair<Rule>) -> Option<String> {
    let direct = |p: &Pair<Rule>| {
        p.clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::id)
            .map(|c| c.as_str().to_string())
    };
    if let Some(id) = direct(pair) {
        return Some(id);
    }
    if pair.as_rule() == Rule::import_field {
        // `(import "mod" "name" (memory $foo 1))` — the id belongs to the
        // descriptor, exactly one level in.
        return pair.clone().into_inner().find_map(|c| direct(&c));
    }
    None
}

fn validate_module(pair: &Pair<Rule>) -> Result<(), String> {
    use std::collections::HashSet;
    let mut export_names: HashSet<String> = HashSet::new();
    let mut func_names: HashSet<String> = HashSet::new();
    let mut ids_by_space: std::collections::HashMap<&'static str, HashSet<String>> =
        std::collections::HashMap::new();
    let mut func_count: usize = 0;
    let mut immut_globals: HashSet<String> = HashSet::new();
    let mut start_target: Option<String> = None;
    let mut start_count = 0usize;
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
        // The id, when the field declares one, belongs to that space alone.
        // An import claims a name in the same space as a definition of its
        // kind, so both are recorded here.
        let space = match inner.as_rule() {
            Rule::func_field => Some("func"),
            Rule::table_field => Some("table"),
            Rule::memory_field => Some("memory"),
            Rule::global_field => Some("global"),
            Rule::tag_field => Some("tag"),
            Rule::import_field => {
                let text = inner.as_str();
                // The imported kind is the inner descriptor.
                ["func", "table", "memory", "global", "tag"]
                    .iter()
                    .find(|kind| text.contains(&format!("({kind}")))
                    .copied()
            }
            _ => None,
        };
        if let Some(space) = space {
            if let Some(id) = find_first_id(&inner) {
                // `$""` names nothing: an identifier's characters are its
                // name, and the empty name is not one (`id.wast`).
                if id == "$\"\"" || id == "$" {
                    return Err("empty identifier".to_string());
                }
                if !ids_by_space.entry(space).or_default().insert(id.clone()) {
                    return Err(format!("duplicate {space}: {id}"));
                }
            }
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
                start_count += 1;
                if let Some(idx) = inner.into_inner().find(|c| c.as_rule() == Rule::index) {
                    start_target = Some(idx.as_str().to_string());
                }
            }
            _ => {}
        }
    }

    // A module has AT MOST ONE start section (`start.wast`: "multiple start
    // sections"). A single `start_target` Option hid this — the second field
    // simply overwrote the first — and the spec calls the text malformed, not
    // invalid, so it has to be caught here rather than by a validator.
    if start_count > 1 {
        return Err("multiple start sections".to_string());
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

fn walk_module(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    validate_module(&pair)?;
    let span = to_span(&pair);
    // A wast script may declare SEVERAL modules; each is a distinct instance,
    // so each needs its own class name. Sharing one name made every class after
    // the first shadow its predecessors at hoist time, and an `invoke` of an
    // earlier module's export resolved against the LAST class → "undefined is
    // not callable". Numbering starts at the SECOND module so a single-module
    // script keeps the plain `__wasm_module` name (which the compiler's passive
    // element-segment directive resolves against by that literal name).
    let module_seq = {
        let n = __w.module_seq;
        __w.module_seq = n + 1;
        n
    };
    __w.current_module_seq = module_seq;
    let default_class_name = if module_seq == 0 {
        "__wasm_module".to_string()
    } else {
        format!("__wasm_module_{module_seq}")
    };
    // `(module definition …)` is DECLARED, never instantiated — the `(module
    // instance $I $M)` form instantiates it later. Its class and all its
    // prescan registries are still published (that is what a future instance
    // needs), but every INSTANTIATION EFFECT is suppressed below: declared
    // memories and tables, data and element segments, tags, and the start
    // function. Walking one as an ordinary module is what made table.wast
    // allocate ~70GB and memory.wast 4GiB/20s for modules that never run.
    let is_definition = pair
        .clone()
        .into_inner()
        .any(|c| c.as_rule() == Rule::module_definition_kw);
    // The class name is needed by several prescans below (they publish
    // per-module registries keyed by it), so resolve it once, up front.
    let prescan_class_name = pair
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
        .unwrap_or_else(|| default_class_name.clone());
    let mut module_name: Option<String> = None;
    let mut members: Vec<ClassMember> = Vec::new();
    let mut pre_stmts: Vec<Statement> = Vec::new(); // before class (globals)
    let mut post_stmts: Vec<Statement> = Vec::new(); // after class (start, exports, imports)

    let mut index_arities = Vec::new();
    let mut name_arities = HashMap::new();
    let mut index_results = Vec::new();
    let mut name_results = HashMap::new();
    // Function index → the method name it compiles to. An unnamed function is
    // named after its first inline export, else synthesized from its index —
    // never a shared constant, since a module may hold several unnamed
    // functions and `(export "a" (func 0))` must name exactly one of them.
    let mut index_names: Vec<String> = Vec::new();

    // 1. Pre-scan imports. Params live inside `typeuse` (and, for imports, inside
    //    `import_desc`), so the signature scan must descend, not read direct children.
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::import_field {
                    // ONLY function imports occupy the function index space.
                    // `import_desc` also spells table / memory / global / tag
                    // (grammar.pest), and each of those has its own index
                    // space — pushing them here shifted every function index
                    // by the number of non-func imports ahead of it, so
                    // `(import "m" "g" (global i32))` before a function made
                    // `call 1` miss it by one. Same `(func` test the export
                    // prescan below uses on its own descriptor.
                    let is_func_import = inner
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::import_desc)
                        .is_some_and(|d| d.as_str().trim_start().starts_with("(func"));
                    if !is_func_import {
                        continue;
                    }
                    // The (module, name) pair is right here in the source —
                    // bind the local alias to it so the call can be lowered
                    // from what the import declared rather than from a table.
                    if let (Some(alias), Some((m, n))) =
                        (scan_func_signature(inner.clone()).0, scan_import_names(&inner))
                    {
                        // An instantiation's `with` clause wins over the
                        // import's own (module, name) pair. That is what
                        // `instantiate` MEANS: the module names a slot, the
                        // instantiation says what fills it. Without this a
                        // core module inside a component could only reach a
                        // canon built-in by spelling the canonidx into the
                        // import name itself.
                        //
                        // A `with` clause may supply either a CANON row or
                        // another instance's export. Only the first is a host
                        // callee; the second is an ordinary module-to-module
                        // link and is bound in step 3c against `import_alias`.
                        // Falling through to `host:{m}:{n}` for it would send
                        // the call to a host function that may well exist,
                        // which is a wrong answer rather than a missing one.
                        match __w.component_imports.get(&(m.clone(), n.clone())) {
                            Some(CoreFunc::Canon(builtin, canonidx)) => {
                                __w.host_import_alias
                                    .insert(alias, format!("host:canon:{builtin}@{canonidx}"));
                            }
                            Some(CoreFunc::Module(..)) => {}
                            None => {
                                __w.host_import_alias.insert(alias, format!("host:{m}:{n}"));
                            }
                        }
                    }
                    let (name, params_count, results_count) = scan_func_signature(inner);
                    index_arities.push(params_count);
                    index_results.push(results_count);
                    // Imported functions occupy the LEADING function indices, so
                    // the index→name map must cover them for `(export "e"
                    // (func N))` to name the right entity.
                    index_names.push(
                        name.clone()
                            .unwrap_or_else(|| format!("__wasm_import_{}", index_names.len())),
                    );
                    if let Some(n) = name {
                        name_arities.insert(n.clone(), params_count);
                        name_results.insert(n, results_count);
                    }
                }
            }
        }
    }

    // 2. Pre-scan defined functions
    //
    // ⛔An EXPORT NAME and a `$id` are DIFFERENT NAMESPACES. An unnamed
    // `(func (export "get") …)` is reached in the lowered module class by its
    // export name, and that is fine right up until some OTHER function is
    // declared `$get` — then both want the same method and one silently
    // replaces the other. `gc/array.wast` has exactly that pair, and the
    // exported wrapper's `call $get` resolved to ITSELF: an infinite
    // recursion that surfaced as a stack overflow, not as a name clash.
    //
    // So the ids are collected FIRST, over the whole module: a later
    // declaration is as much of a collision as an earlier one, and the
    // single-pass `defined_names` below cannot see it yet.
    let declared_func_ids: std::collections::HashSet<String> = pair
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::module_field)
        .filter_map(|c| c.into_inner().next())
        .filter(|inner| inner.as_rule() == Rule::func_field)
        .filter_map(|inner| scan_func_signature(inner).0)
        .collect();
    let mut defined_names: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut export_map: HashMap<String, String> = HashMap::new();
    let mut defined_func_count = 0usize;
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
                    // An export name is only usable as the method name when no
                    // function claims it as a `$id` — see the note above.
                    let method = name.clone().or_else(|| {
                        exports
                            .iter()
                            .find(|e| !declared_func_ids.contains(*e))
                            .cloned()
                    });
                    // The name the function is ACTUALLY declared under — the
                    // synthetic one when neither an id nor a usable export name
                    // is available. Every export still routes to it through
                    // `export_map`, so `(invoke "get")` reaches the wrapper even
                    // though the wrapper is not called `get`.
                    let declared = method
                        .clone()
                        .unwrap_or_else(|| format!("__wasm_func_{}", index_names.len()));
                    index_names.push(declared.clone());
                    defined_func_count += 1;
                    if method.is_some() {
                        // An unnamed exported func is reached by its export name
                        // (e.g. `_start`); key its signature there too so the
                        // entry auto-invoke can tell whether it yields a value.
                        name_results.entry(declared.clone()).or_insert(results_count);
                        name_arities.entry(declared.clone()).or_insert(params_count);
                    }
                    for e in exports {
                        export_map.insert(e, declared.clone());
                    }
                } else if inner.as_rule() == Rule::export_field {
                    // `(export "e" (func $g))` / `(export "e" (func 0))`: map the
                    // export name to the function's method name. The target sits
                    // inside `export_desc → index`, NOT as a direct child, and a
                    // numeric index names the function positionally. Only `func`
                    // descriptors belong in this map — a table/memory/global
                    // export names a different entity, and mapping it here would
                    // make `invoke` call a non-function.
                    let mut ename: Option<String> = None;
                    let mut target: Option<String> = None;
                    for c in inner.into_inner() {
                        match c.as_rule() {
                            Rule::string => ename = Some(unquote(c.as_str())),
                            Rule::export_desc => {
                                if !c.as_str().trim_start().starts_with("(func") {
                                    continue;
                                }
                                if let Some(idx) =
                                    c.into_inner().find(|p| p.as_rule() == Rule::index)
                                {
                                    target = resolve_func_index_name(&idx, &index_names);
                                }
                            }
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

    __w.func_index_arities = index_arities;
    __w.func_name_arities = name_arities;
    __w.func_index_results = index_results;
    __w.func_name_results = name_results;
    __w.defined_func_names = defined_names;
    __w.export_func_map = export_map;
    __w.func_index_name = index_names.clone();

    // 3. Pre-scan struct type definitions to know field counts for struct.new arity
    let mut struct_counts: HashMap<String, usize> = HashMap::new();
    let mut func_param_counts: HashMap<String, usize> = HashMap::new();
    // type name → (param val types, result val types) — the structural
    // identity `ref.test (ref $t)` on a function reference matches against.
    let mut func_sigs: HashMap<String, (Vec<String>, Vec<String>)> = HashMap::new();
    let mut func_parents: HashMap<String, String> = HashMap::new();
    let mut func_result_counts: HashMap<String, usize> = HashMap::new();
    let mut array_elem_types: HashMap<String, String> = HashMap::new();
    // Every declared type's name, in order, so a numeric parent index resolves
    // to a name. GC structs collected here (name, raw parent ref, field count).
    let mut type_order: Vec<String> = Vec::new();
    // (qualified name, raw parent ref, composite text) in declaration order.
    let mut type_shapes: Vec<(String, Option<String>, String)> = Vec::new();
    let mut struct_types_raw: Vec<(String, Option<String>, usize)> = Vec::new();
    let mut struct_field_types_map: HashMap<String, Vec<String>> = HashMap::new();
    let mut struct_field_ids_map: HashMap<String, Vec<String>> = HashMap::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::type_field {
                    // The composite type's own TEXT, for canonicalisation
                    // below. Taken verbatim (whitespace-normalised) rather than
                    // rebuilt from the parsed shape, because the parsed shape
                    // drops FIELD MUTABILITY — `array_elem_type` descends
                    // through `(mut i32)` and answers "i32" — and merging a
                    // mutable type with an immutable one would be a wrong
                    // answer, not a missed optimisation.
                    let comp_text = composite_type_text(&inner);
                    let mut type_name: Option<String> = None;
                    let mut field_count = 0usize;
                    let mut field_types: Vec<String> = Vec::new();
                    let mut field_ids: Vec<String> = Vec::new();
                    let mut is_struct = false;
                    let mut func_params: Option<usize> = None;
                    let mut func_results: Option<usize> = None;
                    let mut func_sig: Option<(Vec<String>, Vec<String>)> = None;
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
                                        field_ids = struct_field_names(&inner2);
                                        field_count = field_types.len();
                                    }
                                    // ⚠ A FUNC TYPE CAN BE `sub`-WRAPPED TOO.
                                    //
                                    // This branch handled `array` and `struct`
                                    // inside `(sub … composite)` but not `func`,
                                    // so `(type $t (sub (func (result funcref))))`
                                    // recorded NO signature — only the bare
                                    // `(type $t (func …))` spelling did. The
                                    // consequence was silent: `func_field_signature`
                                    // falls back to `type_func_sigs` for a
                                    // `(func $f (type $t) …)` declaration, found
                                    // nothing, and registered an EMPTY signature,
                                    // so `ref.test (ref null $t) (ref.func $f)`
                                    // compared `[]→[]` against `[]→[funcref]` and
                                    // answered 0 for a function of exactly that
                                    // type.
                                    if inner2.as_rule() == Rule::func_type {
                                        func_sig = Some(func_type_signature(&inner2));
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
                                        field_ids = struct_field_names(&inner2);
                                            field_count = field_types.len();
                                        }
                                        // `(struct_subtype field* $Base)` — legacy
                                        // GC-MVP form: fields then the supertype.
                                        Rule::struct_subtype => {
                                            is_struct = true;
                                            field_types = struct_field_types(&inner2);
                                        field_ids = struct_field_names(&inner2);
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
                                            // The TYPES, for structural matching
                                            // (`Comptype_sub/func`), alongside the
                                            // counts the rest of the walker uses.
                                            func_sig = Some(func_type_signature(&inner2));
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
                    //
                    // ⚠ QUALIFIED BY MODULE. Each wast module has its OWN type
                    // index space, but the compiler keeps ONE script-wide type
                    // table keyed by name — so `$t` declared in two modules was
                    // one row, and the second declaration MUTATED the first.
                    // Measured: `(type $t (sub (struct)))` in module 1 and
                    // `(type $t (sub (func …)))` in module 3 left module 1's
                    // struct row with `kind = Func`, and its casts then trapped.
                    // A wast module cannot name another module's type, so
                    // qualifying is exactly the semantics.
                    let name = qualify_type_name(
                        __w,
                        &type_name
                            .clone()
                            .unwrap_or_else(|| format!("__wast_type_{}", type_order.len())),
                    );
                    type_order.push(name.clone());
                    type_shapes.push((name.clone(), parent_ref.clone(), comp_text));
                    if is_struct {
                        struct_counts.insert(name.clone(), field_count);
                        struct_types_raw.push((name.clone(), parent_ref.clone(), field_count));
                        struct_field_types_map.insert(name.clone(), field_types.clone());
                        struct_field_ids_map.insert(name.clone(), field_ids.clone());
                    }
                    if let Some(r) = func_results {
                        func_result_counts.insert(name.clone(), r);
                    }
                    if let Some(n) = func_params {
                        func_param_counts.insert(name.clone(), n);
                    }
                    if let Some(sig) = func_sig {
                        func_sigs.insert(name.clone(), sig);
                        // A func type's supertype was dropped here: only
                        // `is_struct` rows reached `struct_types_raw`, so
                        // `(type $sub (sub $super (func …)))` registered with no
                        // parent and became indistinguishable from `$super`.
                        if let Some(p) = parent_ref.clone() {
                            func_parents.insert(name.clone(), qualify_type_name(__w, &p));
                        }
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
                    // `type_order` already holds QUALIFIED names.
                    type_order.get(i).cloned()
                } else {
                    // ⚠ A NAMED parent needs the same module qualifier its
                    // declaration got. Qualifying the child and leaving
                    // `$Base` bare pointed the subtype edge at a row that does
                    // not exist, so `(type $Sub (struct_subtype … $Base))` lost
                    // its supertype and `ref.cast` to `$Sub` failed on a value
                    // that really was one.
                    Some(qualify_type_name(__w, &p))
                }
            });
            (name, parent, fields)
        })
        .collect();
    __w.struct_types = struct_types;
    __w.struct_field_types = struct_field_types_map;
    __w.struct_field_ids = struct_field_ids_map;

    // ── Type CANONICALISATION ────────────────────────────────────────────
    //
    // ⚠ WASM 3.0 type identity is STRUCTURAL, not nominal. Two separately
    // declared types with the same supertype and the same composite shape are
    // THE SAME TYPE, so a value made with one casts to the other:
    //
    //     (type $t1  (sub $t0 (struct (field i32))))
    //     (type $t1' (sub $t0 (struct (field i32))))
    //     (ref.cast (ref $t1') (… a $t1 …))     ;; must SUCCEED
    //
    // We compared names, through a nominal subtype graph in which `$t1'` is
    // neither a subtype nor a supertype of `$t1`, so the cast trapped with
    // "value is not m#1#t1'" — a message that reads as a correct rejection.
    // That is `gc/ref_cast`, `gc/ref_test`, `gc/type-subtyping` and
    // `gc/br_on_cast_fail`, every one of them in its `test-canon` function.
    //
    // The key is the CANONICALISED parent plus the composite text, so a chain
    // canonicalises from the root down; declaration order is enough for the
    // non-recursive case, which is what the suite's `test-canon` exercises.
    // ⛔Scoped to ONE MODULE on purpose. The spec canonicalises across the
    // whole store, but our type names are module-qualified precisely so one
    // module cannot disturb another's rows, and merging across that boundary
    // would undo it for a case the suite does not test.
    {
        let mut canonical: HashMap<String, String> = HashMap::new();
        let mut by_shape: HashMap<String, String> = HashMap::new();
        let mut rec_shape: HashMap<String, String> = HashMap::new();
        // ⛔ A REC GROUP IS PART OF THE IDENTITY. Under iso-recursive
        // equivalence a type declared inside `(rec …)` is identified by its
        // WHOLE GROUP plus its POSITION in it — not by its own shape. Keying on
        // the shape alone merged `$f1` and `$f2` in all three of
        // `type-rec.wast`'s dynamic-matching modules, when only the FIRST (two
        // structurally identical groups) is actually the same type; the other
        // two differ by member ORDER and by group SIZE respectively, and both
        // must trap on `call_indirect`.
        //
        // A type outside any `(rec …)` is its own singleton group, so its key
        // gains a constant prefix and identical standalone types still merge —
        // which is what `test-canon` in the gc suite pins.
        let rec_of = rec_groups_of_module(&pair).unwrap_or_default();
        let group_of = |i: usize| rec_of.get(i).copied().unwrap_or(usize::MAX - i);
        let members_of = |g: usize| -> Vec<usize> {
            (0..type_shapes.len()).filter(|j| group_of(*j) == g).collect()
        };
        // ⛔ A REFERENCE INTO ONE'S OWN GROUP IS A POSITION, NOT A NAME. Under
        // iso-recursive equivalence a group is compared after unrolling, so
        // `(rec (type $t1 (func (result (ref null $t1)))))` and the same
        // written with `$t2` are ONE type — the recursive reference points at
        // "member 0 of this group" in both. Comparing the source text keeps the
        // spelling and splits them, which trapped a valid `call_indirect`.
        //
        // References OUT of the group stay nominal (qualified): resolving those
        // to a canonical form would need a fixpoint over the whole module, and
        // splitting there costs a merge, not a wrong trap.
        let positional_text = |i: usize| -> String {
            let members = members_of(group_of(i));
            let text = &type_shapes[i].2;
            if !text.contains('$') {
                return text.clone();
            }
            let mut out = String::with_capacity(text.len());
            let mut rest = text.as_str();
            while let Some(at) = rest.find('$') {
                out.push_str(&rest[..at]);
                let tail = &rest[at + 1..];
                let end = tail
                    .find(|c: char| !(c.is_alphanumeric() || "_.+-*/\\^~=<>!?@#$%&|:'`".contains(c)))
                    .unwrap_or(tail.len());
                if end == 0 {
                    out.push('$');
                    rest = tail;
                    continue;
                }
                let qualified = qualify_type_name(__w, &tail[..end]);
                match members.iter().position(|j| type_shapes[*j].0 == qualified) {
                    Some(pos) => out.push_str(&format!("rec.{pos}")),
                    None => out.push_str(&qualified),
                }
                rest = &tail[end..];
            }
            out.push_str(rest);
            out
        };
        let shapes: Vec<String> = (0..type_shapes.len()).map(positional_text).collect();
        let group_shape = |g: usize| -> (String, Vec<usize>) {
            let members = members_of(g);
            let text = members
                .iter()
                .map(|j| shapes[*j].clone())
                .collect::<Vec<_>>()
                .join(";");
            (text, members)
        };
        for (i, (name, parent_ref, _)) in type_shapes.iter().enumerate() {
            let parent = parent_ref.as_ref().map(|p| {
                let resolved = match p.parse::<usize>() {
                    Ok(i) => type_order.get(i).cloned().unwrap_or_else(|| p.clone()),
                    Err(_) => qualify_type_name(__w, p),
                };
                canonical.get(&resolved).cloned().unwrap_or(resolved)
            });
            let (gtext, members) = group_shape(group_of(i));
            let pos = members.iter().position(|j| *j == i).unwrap_or(0);
            let key = format!(
                "{}#{}#{}|{}",
                gtext,
                pos,
                parent.unwrap_or_default(),
                shapes[i]
            );
            let winner = by_shape.entry(key).or_insert_with(|| name.clone()).clone();
            rec_shape.insert(winner.clone(), format!("{}:{}", members.len(), pos));
            canonical.insert(name.clone(), winner);
        }
        // A numeric type reference resolves through `type_index_name`, so that
        // list has to carry the canonical names too — otherwise `(type 3)` and
        // `(type $t3)` would name different rows for the same declaration.
        __w.type_index_name = type_order
            .iter()
            .map(|n| canonical.get(n).cloned().unwrap_or_else(|| n.clone()))
            .collect();
        __w.type_canonical = canonical;
        // Script-wide: names are module-qualified, so one map serves every
        // module and a later module cannot overwrite an earlier one's row.
        __w.type_rec_shape.extend(rec_shape);
    }
    __w.struct_field_counts = struct_counts;
    __w.type_func_params = func_param_counts;

    // Call Tags proposal: record every `(call_tag …)`'s arity BEFORE any body
    // is folded. A tag may be referenced before its declaration appears — wat
    // permits forward references — and the folder needs the arity to know how
    // many operands `call_with_tag` takes. Scanned here for the same reason
    // `type_func_params` is: by the time instructions fold, it must be known.
    {
        let mut scanned: Vec<(String, (usize, usize))> = Vec::new();
        for field in pair.clone().into_inner() {
            for inner in field.into_inner() {
                if inner.as_rule() != Rule::call_tag_field {
                    continue;
                }
                let mut name: Option<String> = None;
                let mut shape = (0usize, 0usize);
                for child in inner.into_inner() {
                    match child.as_rule() {
                        Rule::id => name = Some(child.as_str().to_string()),
                        Rule::typeuse => shape = count_typeuse_params_results(__w, &child),
                        _ => {}
                    }
                }
                if let Some(n) = name {
                    scanned.push((n, shape));
                }
            }
        }
        for (n, shape) in scanned {
            __w.call_tag_params.insert(n, shape);
        }
        // An IMPORTED call tag's local alias must be known before any body is
        // walked, or a `call_with_tag $local` inside a func is folded before
        // the import statement registers the alias and names a tag that does
        // not exist.
        let mut aliases: Vec<(String, String, (usize, usize))> = Vec::new();
        for field in pair.clone().into_inner() {
            for inner in field.into_inner() {
                if inner.as_rule() != Rule::import_field {
                    continue;
                }
                let strings: Vec<String> = inner
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::string)
                    .map(|c| unquote(c.as_str()))
                    .collect();
                let Some(desc) = inner
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_desc)
                else {
                    continue;
                };
                if !desc.as_str().trim_start().starts_with("(call_tag") {
                    continue;
                }
                let mut local = String::new();
                let mut shape = (0usize, 0usize);
                for child in desc.into_inner() {
                    match child.as_rule() {
                        Rule::id => {
                            local = child.as_str().trim_start_matches('$').to_string()
                        }
                        Rule::typeuse => shape = count_typeuse_params_results(__w, &child),
                        _ => {}
                    }
                }
                if strings.len() >= 2 && !local.is_empty() {
                    aliases.push((local, strings[1].clone(), shape));
                }
            }
        }
        for (local, external, shape) in aliases {
            __w.call_tag_params.insert(local.clone(), shape);
            __w.call_tag_params.insert(external.clone(), shape);
            __w.call_tag_alias.insert(local, external);
        }
    }
    __w.type_func_results = func_result_counts;
    // ⛔ KEYED PRE-CANONICAL, READ POST-CANONICAL. These keys were built by
    // `qualify_type_name` during the pre-scan, when `type_canonical` is still
    // EMPTY — that emptiness is deliberate, it is what lets the pass read raw
    // names. But `type_canonical` is populated further up (`__w.type_canonical
    // = canonical`) BEFORE anything looks a signature up, and every reader goes
    // through the same funnel, which now answers with the CANONICAL name. So a
    // type with a canonical form was stored under one key and asked for under
    // another, and `(type $t)` resolved to nothing — which
    // `typeuse_signature` then formatted as `"->"`, a real signature meaning
    // "no params, no results".
    //
    // Re-key through the mapping now that it exists, so both sides agree.
    // Two names collapsing onto one canonical key is correct: they ARE the
    // same type.
    // The VALUES need the same treatment as the keys: a stored signature holds
    // `(ref $s1)` verbatim from the pre-scan, and the call site it is compared
    // against resolves its own reference canonically — so two equal types read
    // as different signatures.
    __w.type_func_sigs = func_sigs
        .into_iter()
        .map(|(k, (ps, rs))| {
            let canon_all = |v: Vec<String>| {
                v.into_iter().map(|t| canonical_val_type(__w, &t)).collect::<Vec<_>>()
            };
            (
                __w.type_canonical.get(&k).cloned().unwrap_or(k),
                (canon_all(ps), canon_all(rs)),
            )
        })
        .collect();
    __w.type_func_parent = func_parents;
    // Each DEFINED function's own signature, collected here because `pair` is
    // consumed before the directives are emitted. Resolving a `(type $t)`
    // reference needs `type_func_sigs` above, so this must come after it.
    let defined_func_sigs: Vec<(String, (Vec<String>, Vec<String>), Option<String>)> = pair
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::module_field)
        .filter_map(|c| c.into_inner().next())
        .filter(|f| f.as_rule() == Rule::func_field)
        .enumerate()
        .filter_map(|(fi, f)| {
            // ⚠ AN ANONYMOUS FUNCTION STILL HAS A SIGNATURE.
            //
            // This used to `?` out when there was no `$id`, so
            // `(func (export "f") (type $f))` registered NOTHING — no
            // structural signature and no declared type — and `ref.test`
            // against a concrete `(ref $t)` on it answered 0. The naming site
            // falls back to the first EXPORT name and then to the declaration
            // ordinal; this mirrors that, or the two disagree about which
            // function is which.
            let name = f
                .clone()
                .into_inner()
                .find(|c| c.as_rule() == Rule::id)
                .map(|c| c.as_str()[1..].to_string())
                .or_else(|| {
                    f.clone()
                        .into_inner()
                        .filter(|c| c.as_rule() == Rule::export_inline)
                        .filter_map(|c| c.into_inner().find(|p| p.as_rule() == Rule::string))
                        .map(|s| unquote(s.as_str()))
                        .next()
                })
                .unwrap_or_else(|| format!("__wasm_func_{fi}"));
            Some((
                name,
                func_field_signature(__w, &f),
                func_field_declared_type(__w, &f),
            ))
        })
        .collect();
    __w.array_elem_type = array_elem_types;
    // Per-MODULE counter; `elem_index_base` carries the script-wide offset it
    // is added to, so the two together number segments the way the compiler
    // stores them.
    __w.elem_seg_counter = 0;
    // Pre-scan element segments, so `(elem $e …)` resolves and a numeric
    // elemidx shifts by the base. Declarative and active segments occupy a slot
    // too — the index space counts every `(elem …)` in declaration order.
    {
        let mut elem_names: HashMap<String, usize> = HashMap::new();
        let mut elem_idx = 0usize;
        let elem_base = __w.elem_index_base;
        for child in pair.clone().into_inner() {
            if child.as_rule() == Rule::module_field {
                if let Some(inner) = child.into_inner().next() {
                    if inner.as_rule() == Rule::elem_field {
                        if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                            elem_names.insert(id.as_str()[1..].to_string(), elem_base + elem_idx);
                        }
                        elem_idx += 1;
                    }
                }
            }
        }
        __w.elem_name_index = elem_names;
    }

    // 3a. Pre-scan tables so named tables (`$t1`) resolve to their declaration
    //     index for `elem` population and `call_indirect $t` dispatch.
    let mut table_names: HashMap<String, usize> = HashMap::new();
    let mut table_idx = 0usize;
    let table_base = __w.table_index_base;
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::table_field {
                    if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                        table_names.insert(id.as_str()[1..].to_string(), table_base + table_idx);
                    }
                    table_idx += 1;
                }
            }
        }
    }
    __w.table_name_index = table_names;
    // Advance the base only after this module is fully walked (deferred to the
    // end of this function); record the count here.
    let module_table_count = table_idx;

    // 3a'. Pre-scan memories.
    //
    // ⚠ A memidx is MODULE-RELATIVE and a memory is not always NEW. An
    // IMPORTED memory occupies an index in this module while BEING the
    // exporter's memory, so a per-module base cannot express it — `memory_slots`
    // maps this module's index space onto the script's, one entry per declared
    // memory, and an import's entry points back at the memory it aliases.
    //
    // Before this, `import_inline` was in the grammar and read NOWHERE: the
    // import was silently dropped and a fresh empty memory declared in its
    // place. Both modules then had their own bytes, which is why
    // `multi-memory/load1`, `data0`, `linking1`, `imports1` and friends wrote
    // through the import and read nothing back through the exporter, and why
    // `memory_size_import` answered the IMPORT's declared minimum instead of
    // the exporter's actual size.
    //
    // ⚠ An import's limits are a LINK-TIME CONSTRAINT, not a sizing fact —
    // `memory_size_import` imports `1 5` from a memory that has `2 4` and must
    // answer 2. Aliasing therefore carries no limits at all.
    //
    // Two spellings declare the same thing and both count: the abbreviated
    // `(memory $m (import "M" "mem") …)` and the standalone
    // `(import "M" "mem" (memory …))`. Missing either one would misalign every
    // later numeric memidx in the module — strictly worse than a duplicate
    // memory, since it would silently name a DIFFERENT existing one.
    let mut memory_names: HashMap<String, usize> = HashMap::new();
    let mut memory_slots: Vec<usize> = Vec::new();
    let mut memory_field_info: Vec<(usize, bool)> = Vec::new();
    let mut memory_exports: HashMap<String, usize> = HashMap::new();
    let mut defined_memories = 0usize;
    let memory_base = __w.memory_index_base;
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        let inner_rule = inner.as_rule();
        // `(memory …)`, or an `(import … (memory …))` whose descriptor is one.
        let (decl, import_pair) = match inner_rule {
            Rule::memory_field => {
                let imp = inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_inline);
                (Some(inner.clone()), imp)
            }
            Rule::import_field => {
                let desc = inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_desc)
                    .filter(|d| {
                        d.as_str()
                            .trim_start_matches('(')
                            .trim_start()
                            .starts_with("memory")
                    });
                match desc {
                    Some(d) => (Some(d), Some(inner.clone())),
                    None => (None, None),
                }
            }
            _ => (None, None),
        };
        let Some(decl) = decl else { continue };
        // The two module/name strings sit on `import_inline` directly, and on
        // `import_field` ahead of its descriptor.
        let aliased = import_pair.and_then(|imp| {
            let strings: Vec<String> = imp
                .into_inner()
                .filter(|c| c.as_rule() == Rule::string)
                .map(|s| unquote(s.as_str()))
                .take(2)
                .collect();
            let (m, n) = (strings.first()?, strings.get(1)?);
            let class = __w.registered_module_class.get(m)?;
            __w.module_memory_exports.get(class)?.get(n).copied()
        });
        // An import we cannot resolve — `"spectest"`'s host memory, or a module
        // that was never registered — has no memory to alias.
        //
        // ⚠ In the STANDALONE `(import … (memory …))` spelling such an import
        // declares nothing at all today (`walk_import_field` emits a noop), so
        // giving it an index here would put a slot in the space with no memory
        // behind it and shift every later module's base. It is skipped, which
        // leaves those modules numbered exactly as they were — a known
        // deviation from the spec's "imports occupy the low indices", recorded
        // rather than half-fixed. The INLINE `(memory $m (import …) 1)`
        // spelling does declare a memory, so it takes a real slot.
        let is_standalone = matches!(inner_rule, Rule::import_field);
        let slot = match aliased {
            Some(s) => s,
            None if is_standalone => continue,
            None => {
                let s = memory_base + defined_memories;
                defined_memories += 1;
                s
            }
        };
        if let Some(id) = decl
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::id)
        {
            memory_names.insert(id.as_str()[1..].to_string(), slot);
        }
        for e in decl
            .clone()
            .into_inner()
            .filter(|c| c.as_rule() == Rule::export_inline)
        {
            if let Some(s) = e.into_inner().find(|c| c.as_rule() == Rule::string) {
                memory_exports.insert(unquote(s.as_str()), slot);
            }
        }
        memory_slots.push(slot);
        // `walk_memory_field` is driven by a counter over `memory_field`s only,
        // so it needs its own parallel record: which slot each one owns, and
        // whether it ALIASES (in which case it must declare nothing).
        if matches!(inner_rule, Rule::memory_field) {
            memory_field_info.push((slot, aliased.is_some()));
        }
    }
    // `(export "name" (memory $m))` names an already-declared memory.
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        if inner.as_rule() != Rule::export_field {
            continue;
        }
        let name = inner
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::string)
            .map(|s| unquote(s.as_str()));
        let target = inner
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::export_desc)
            .filter(|d| {
                d.as_str()
                    .trim_start_matches('(')
                    .trim_start()
                    .starts_with("memory")
            })
            .and_then(|d| d.into_inner().find(|c| c.as_rule() == Rule::index));
        if let (Some(name), Some(idx)) = (name, target) {
            let t = idx.as_str().trim();
            let slot = match t.strip_prefix('$') {
                Some(id) => memory_names.get(id).copied(),
                None => t.parse::<usize>().ok().and_then(|n| memory_slots.get(n).copied()),
            };
            if let Some(slot) = slot {
                memory_exports.insert(name, slot);
            }
        }
    }
    __w.memory_name_index = memory_names;
    __w.memory_slots = memory_slots;
    __w.memory_field_info = memory_field_info;
    __w.module_memory_exports
        .insert(prescan_class_name.clone(), memory_exports);
    // Advance the base for the NEXT module only after this one is fully
    // walked (deferred to the end of this function); record the count here.
    // ⚠ DEFINED memories only — an aliased import allocates nothing.
    let module_memory_count = defined_memories;

    // 3a'''. Pre-scan data segments, for the same reason and in the same index
    //        space discipline as memories: `memory.init`/`data.drop` name a
    //        segment module-relatively, the compiler stores them all in one
    //        script-wide list. ACTIVE segments count too — WASM gives active and
    //        passive segments a single index space, and so does the compiler.
    let mut data_names: HashMap<String, usize> = HashMap::new();
    let mut data_idx = 0usize;
    let data_base = __w.data_index_base;
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::data_field {
                    if let Some(id) = inner.into_inner().find(|c| c.as_rule() == Rule::id) {
                        data_names.insert(id.as_str()[1..].to_string(), data_base + data_idx);
                    }
                    data_idx += 1;
                }
            }
        }
    }
    __w.data_name_index = data_names;
    let module_data_count = data_idx;

    // 3a''. Pre-scan globals so a `global.get N` / `global.set N` by numeric
    //       index resolves to the right binding (each global's `$id`, or a
    //       synthetic `__wasm_global_<i>` when unnamed).
    let mut global_names: Vec<String> = Vec::new();
    let mut global_export_map: HashMap<String, String> = HashMap::new();
    // ⚠ TWO COUNTERS, ON PURPOSE.
    //
    // `global_names.len()` is the INDEX SPACE — imports included, because the
    // spec numbers them first and `global.get <n>` has to agree.
    // `defined_ordinal` counts only DEFINED globals, and is what names an
    // anonymous one, because `walk_global_field` numbers its own fallback the
    // same way and never sees the imports.
    //
    // Collapsing them silently breaks the export map: once a standalone global
    // import occupied a slot, an anonymous `(global (export "h") …)` was NAMED
    // `__wasm_global_m_1` here but DECLARED `__wasm_global_m_0` there, so the
    // export resolved to a binding that was never declared and read undefined.
    let mut defined_ordinal: usize = 0;
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::global_field {
                    let idx = defined_ordinal;
                    defined_ordinal += 1;
                    let binding = global_binding_name(__w, &inner, idx);
                    // Inline exports: `(global $a (export "a") i32 …)`.
                    for e in inner
                        .clone()
                        .into_inner()
                        .filter(|c| c.as_rule() == Rule::export_inline)
                        .filter_map(|c| c.into_inner().find(|p| p.as_rule() == Rule::string))
                    {
                        global_export_map.insert(unquote(e.as_str()), binding.clone());
                    }
                    // `(global $g (import "m" "e") …)` is a second NAME for the
                    // exporting module's global, not a new cell. Resolve it to
                    // the exporter's binding and record the alias, so both the
                    // named and the index form read and write the SAME binding
                    // — a copy would satisfy an immutable import but break a
                    // mutable one, where the spec shares one cell.
                    //
                    // Without this the import was ignored entirely: the field
                    // has no `instr` child, so `walk_global_field` left the
                    // initialiser at its `Expression::int(0)` default and every
                    // read of an imported global answered 0.
                    let imported = inner
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::import_inline)
                        .and_then(|c| scan_import_names(&c))
                        .and_then(|(m, e)| {
                            __w.registered_module_class.get(&m).cloned().and_then(|cls| {
                                __w.module_global_exports
                                    .get(&cls)
                                    .and_then(|ex| ex.get(&e).cloned())
                            })
                        });
                    match imported {
                        Some(exporter_binding) => {
                            __w.global_import_alias
                                .insert(binding.clone(), exporter_binding.clone());
                            // The index form must reach the same cell.
                            global_names.push(exporter_binding);
                        }
                        None => global_names.push(binding),
                    }
                } else if inner.as_rule() == Rule::import_field {
                    // ⚠ THE STANDALONE SPELLING COUNTS TOO.
                    //
                    // `(import "m" "e" (global $g …))` is the same entity as
                    // `(global $g (import "m" "e") …)` above, but it parses as
                    // an `import_field`, so this loop skipped it entirely: no
                    // alias — every read got a fresh binding instead of the
                    // exporter's cell — and no slot, so every LATER global's
                    // numeric index was off by one.
                    //
                    // Identity is what makes it observable. The proposal's
                    // `ref_get_desc.wast` imports a descriptor global and
                    // compares it with `ref.eq` against the descriptor read
                    // back off an object the exporting module allocated; with
                    // a copy the two are different references and the
                    // assertion reads 0 for 1. Nothing about that is specific
                    // to descriptors — a plain `(ref null $s)` global behaves
                    // the same, which is how this was isolated.
                    let Some(desc) = inner
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::import_desc)
                        .filter(|d| d.as_str().trim_start().starts_with("(global"))
                    else {
                        continue;
                    };
                    // An IMPORT is not a defined global, so it does not consume
                    // a `defined_ordinal` — but it does take an index-space
                    // slot, which is why it still pushes to `global_names`.
                    // The fallback name is only reached by an anonymous import,
                    // which cannot be referred to by name anyway.
                    let binding = global_binding_name(__w, &desc, global_names.len());
                    let imported = scan_import_names(&inner).and_then(|(m, e)| {
                        __w.registered_module_class.get(&m).cloned().and_then(|cls| {
                            __w.module_global_exports
                                .get(&cls)
                                .and_then(|ex| ex.get(&e).cloned())
                        })
                    });
                    match imported {
                        Some(exporter_binding) => {
                            __w.global_import_alias
                                .insert(binding.clone(), exporter_binding.clone());
                            global_names.push(exporter_binding);
                        }
                        None => global_names.push(binding),
                    }
                }
            }
        }
    }
    // Standalone `(export "e" (global $g))` / `(export "e" (global 0))`. Walked
    // after the names above so a forward reference resolves.
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::module_field {
            if let Some(inner) = child.into_inner().next() {
                if inner.as_rule() == Rule::export_field {
                    let mut ename: Option<String> = None;
                    let mut target: Option<String> = None;
                    for c in inner.into_inner() {
                        match c.as_rule() {
                            Rule::string => ename = Some(unquote(c.as_str())),
                            Rule::export_desc => {
                                if !c.as_str().trim_start().starts_with("(global") {
                                    continue;
                                }
                                if let Some(idx) =
                                    c.into_inner().find(|p| p.as_rule() == Rule::index)
                                {
                                    target = resolve_func_index_name(&idx, &global_names);
                                }
                            }
                            _ => {}
                        }
                    }
                    if let (Some(e), Some(t)) = (ename, target) {
                        global_export_map.insert(e, t);
                    }
                }
            }
        }
    }
    __w.global_index_name = global_names;
    __w.export_global_map = global_export_map.clone();
    {
        __w.module_global_exports
            .insert(prescan_class_name.clone(), global_export_map)
    };

    // 3b. Pre-scan exception tags so a `catch $e` in any function body knows
    //     the tag's payload arity regardless of source order. Reset first —
    //     the thread-local persists across modules compiled on this thread.
    let module_seq = {
        let cur = __w.module_tag_seq;
        __w.module_tag_seq += 1;
        cur
    };
    let mut tag_arities: HashMap<String, u8> = HashMap::new();
    let mut tag_names: Vec<String> = Vec::new();
    let mut tag_canon: HashMap<String, String> = HashMap::new();
    let mut tag_exports: Vec<(String, String)> = Vec::new();
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        // An IMPORTED tag occupies a slot in the tag index space too. It is
        // spelt two ways — `(tag $x (import "m" "t") …)` and
        // `(import "m" "t" (tag $x …))` — and both must be seen, or every later
        // ordinal shifts and a numeric `throw 1` names the wrong entity.
        let decl = match inner.as_rule() {
            Rule::tag_field => Some(scan_tag_decl(__w, &inner)),
            Rule::import_field => {
                let outer = scan_import_names(&inner);
                inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_desc)
                    .filter(|d| d.as_str().trim_start().starts_with("(tag"))
                    .map(|d| {
                        let mut t = scan_tag_decl(__w, &d);
                        t.3 = outer;
                        t
                    })
            }
            _ => None,
        };
        let Some((name, arity, exports, import)) = decl else {
            continue;
        };
        // An anonymous `(tag (param i32))` is still a tag ENTITY with an index;
        // name it by ordinal so numeric references reach it and two anonymous
        // tags stay distinct.
        let local = name.unwrap_or_else(|| format!("#{}", tag_names.len()));
        // An import is an ALIAS for the exporting module's entity; anything
        // else is this module's own, namespaced so two modules declaring `$e0`
        // stay distinct.
        let canon = match &import {
            Some((m, n)) => __w
                .registered_tags
                .get(&(m.clone(), n.clone()))
                .cloned()
                .unwrap_or_else(|| format!("{m}:{n}")),
            None => format!("m{module_seq}:{local}"),
        };
        for e in exports {
            tag_exports.push((e, canon.clone()));
        }
        tag_arities.insert(canon.clone(), arity);
        tag_canon.insert(local.clone(), canon);
        tag_names.push(local);
    }
    __w.tag_arities = tag_arities;
    __w.tag_index_name = tag_names;
    __w.tag_canon = tag_canon;
    __w.pending_tag_exports = tag_exports;
    __w.tag_decl_ordinal = 0;

    // 3c. Pre-scan FUNCTION imports and bind each local alias to the exporting
    //     module's method. `(register "m")` published that module's class, and
    //     its exports are that class's static methods, so an import is just a
    //     second name for one of them — the whole of wast linking, for
    //     functions. Two spellings again: `(func $f (import "m" "e") …)` and
    //     `(import "m" "e" (func $f …))`.
    let mut import_alias: HashMap<String, (String, String)> = HashMap::new();
    // Counts function imports in declaration order, matching the index
    // pre-scan's counter so an unnamed import gets the same synthetic binding.
    let mut alias_func_import_ordinal = 0usize;
    for child in pair.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        let Some(inner) = child.into_inner().next() else {
            continue;
        };
        let (local, target) = match inner.as_rule() {
            Rule::func_field => {
                let local = inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
                let target = inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_inline)
                    .and_then(|c| scan_import_names(&c));
                (local, target)
            }
            Rule::import_field => {
                let target = scan_import_names(&inner);
                let desc = inner
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::import_desc)
                    .filter(|d| d.as_str().trim_start().starts_with("(func"));
                // An UNNAMED function import still occupies an index, and the
                // index pre-scan names it by ordinal — bind the same name here
                // so `call 0` reaches the exporter's method exactly as
                // `call $named` does.
                let local = desc.map(|d| {
                    let named = d
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::id)
                        .map(|c| c.as_str()[1..].to_string());
                    let name = named
                        .unwrap_or_else(|| imported_func_binding_name(alias_func_import_ordinal));
                    alias_func_import_ordinal += 1;
                    name
                });
                (local, target)
            }
            _ => (None, None),
        };
        let (Some(local), Some((m, e))) = (local, target) else {
            continue;
        };
        // An instantiation's `with` clause wins over the registered-module
        // table: the module names a slot, the instantiation says what fills
        // it. `(register "m")` is the script-level fallback for a module that
        // was NOT instantiated by a component.
        let resolved = match __w.component_imports.get(&(m.clone(), e.clone())) {
            Some(CoreFunc::Module(class, method)) => Some((class.clone(), method.clone())),
            // A canon row is a host callee, already bound above.
            Some(CoreFunc::Canon(..)) => None,
            None => __w.registered_module_class.get(&m).cloned().and_then(|class| {
                __w.module_exports
                    .get(&class)
                    .and_then(|ex| ex.get(&e).cloned())
                    .map(|method| (class, method))
            }),
        };
        // An unresolved import is left alone: it may be a host function the
        // profile's builtin table answers, which is how the WASI tests import.
        if let Some(t) = resolved {
            import_alias.insert(local, t);
        }
    }
    __w.import_alias = import_alias;

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
                // Capture the start function so it can be invoked as a static
                // method of the module class at instantiation.
                //
                // ⛔ `(start 2)` is a NUMERIC function index and `start.wast`
                // writes it that way. Reading only `Rule::id` here left
                // `start_fn_name` None ⇒ no start function was invoked at all
                // ⇒ the module's initialisation silently never ran. Resolve
                // through `resolve_func_index_name`, which is the SHARED answer
                // to "which function does this index name" and already takes
                // both spellings — a private half-answer beside it is how the
                // two disagreed.
                start_fn_name = inner
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::index)
                    .and_then(|idx| resolve_func_index_name(&idx, &__w.func_index_name));
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
    // Publish this module's exports under its class name so a later
    // `(invoke $M "e")` can reach it after other modules have been walked.
    let module_exports = __w.export_func_map.clone();
    {
        __w.module_exports
            .insert(prescan_class_name.clone(), module_exports)
    };
    __w.module_class_name = prescan_class_name;

    let mut global_decl_idx = 0usize;
    // Table population (inline `(elem …)` abbreviation, init expression) targets
    // the SCRIPT's table index space, so it starts at this module's base — not 0.
    let mut table_decl_idx = __w.table_index_base;
    let mut memory_decl_idx = 0usize;
    // Imported functions occupy the leading function indices (prescan step 1),
    // so a defined function's index starts after them.
    let import_func_count = index_names.len() - defined_func_count;
    let mut defined_func_seq = 0usize;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => {
                module_name = Some(child.as_str()[1..].to_string());
            }
            Rule::module_field => {
                let inner = child.into_inner().next().ok_or("Empty module_field")?;
                match inner.as_rule() {
                    Rule::func_field => {
                        // Defined functions follow the imported ones in the
                        // function index space.
                        let idx = import_func_count + defined_func_seq;
                        defined_func_seq += 1;
                        members.push(ClassMember::Method(Box::new(walk_func_field(__w, inner, idx)?)))
                    }
                    // A `func_switch` "has no type and cannot be directly
                    // called", but `ref.func` must still yield a funcref for
                    // it — and in this front end a func IS a member of the
                    // module class, which is what `ref.func` resolves against.
                    // So it gets an empty member to be referenced by; the
                    // dispatch never enters that body, because `call_with_tag`
                    // recognises the chunk as a switch and picks an arm.
                    Rule::func_switch_field => {
                        let sw_name = inner
                            .clone()
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::id)
                            .map(|c| c.as_str().trim_start_matches('$').to_string())
                            .unwrap_or_default();
                        pre_stmts.push(walk_func_switch_field(__w, inner.clone())?);
                        members.push(ClassMember::Method(Box::new(Statement::new(
                            StmtKind::FunctionDecl {
                                name: sw_name,
                                params: Vec::new(),
                                return_type: None,
                                body: Vec::new(),
                                modifiers: Default::default(),
                                handles: Vec::new(),
                                is_async: false,
                                is_generator: false,
                                is_sub: false,
                            },
                        ))));
                    }
                    Rule::import_field => {
                        // An imported tag takes the next slot in the tag index
                        // space, so it must advance the same ordinal counter a
                        // defined `(tag …)` uses to name itself.
                        if inner
                            .clone()
                            .into_inner()
                            .any(|c| {
                                c.as_rule() == Rule::import_desc
                                    && c.as_str().trim_start().starts_with("(tag")
                            })
                        {
                            __w.tag_decl_ordinal += 1;
                        }
                        post_stmts.push(walk_import_field(inner)?)
                    }
                    Rule::export_field => {
                        post_stmts.push(Statement::new(StmtKind::Expr(walk_export_field(inner)?)));
                    }
                    Rule::global_field => {
                        // Globals become top-level let bindings BEFORE the class so that
                        // global.get $name → Ident("name") resolves correctly from methods.
                        // The declaration index gives unnamed globals a stable name that
                        // `global.get <idx>` resolves to (see GLOBAL_INDEX_NAME).
                        // `ref.func $f` lowers to `<ModuleClass>.$f`, and the
                        // class is emitted AFTER the globals — so a global
                        // initialised that way must be declared after it, or
                        // the initialiser reads an undefined class and lands
                        // null. `post_stmts` still runs before `start`, so the
                        // global holds its reference by the time anything can
                        // observe it.
                        let references_class = inner.as_str().contains("ref.func");
                        let (name, init) = walk_global_field(__w, inner, global_decl_idx)?;
                        global_decl_idx += 1;
                        // An IMPORTED global already has a cell — the
                        // exporter's. Declaring a local binding of the same
                        // name would shadow the alias with a fresh 0.
                        if __w.global_import_alias.contains_key(&name) {
                            continue;
                        }
                        let decl = Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(name),
                                type_hint: None,
                                init: Some(init),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        });
                        if references_class {
                            post_stmts.push(decl);
                        } else {
                            pre_stmts.push(decl);
                        }
                    }
                    // `(start $f)` is invoked as a static method at instantiation
                    // in the module assembly below (its name was captured in the
                    // pre-scan); nothing to emit per-field here.
                    Rule::start_field => {}
                    // Linear memory + data segments: emitted before the class so
                    // the compiler lowers them into the script chunk's memory /
                    // data tables (the VM allocates pages and writes active data
                    // at instantiation, before `_start`).
                    Rule::memory_field if !is_definition => {
                        pre_stmts.extend(walk_memory_field(__w, inner, memory_decl_idx)?);
                        memory_decl_idx += 1;
                    }
                    Rule::data_field if !is_definition => pre_stmts.push(walk_data_field(__w, inner)?),
                    Rule::table_field if !is_definition => {
                        let (decl, population) = walk_table_field(__w, inner, table_decl_idx)?;
                        pre_stmts.push(decl);
                        post_stmts.extend(population);
                        table_decl_idx += 1;
                    }
                    // Exception tags: declared before the class so the tag
                    // entity exists in the script chunk; `throw`/`catch` in the
                    // function chunks re-import by name and coalesce to it.
                    // NOT gated on `is_definition`: a tag is an ENTITY the
                    // definition's OWN function bodies reference by name
                    // (`throw`/`catch` re-import and coalesce to it), not an
                    // allocation. instance.wast:7 declares one inside
                    // `(module definition $M …)` and throws it at :44.
                    Rule::tag_field => pre_stmts.push(walk_tag_field(__w, inner)?),
                    // Call Tags proposal: a call tag and a func_switch are
                    // module ENTITIES, like a tag or a global, so they are
                    // declarations that must exist before any body references
                    // them.
                    Rule::call_tag_field => pre_stmts.push(walk_call_tag_field(__w, inner)?),
                    // `(import "m" "e" (call_tag $t …))` — the outer-import
                    // spelling of the same declaration. Both spellings key on
                    // the external name so an import and its export are ONE
                    // tag; identity crossing the boundary is the point.
                    Rule::import_field
                        if inner
                            .clone()
                            .into_inner()
                            .any(|c| c.as_rule() == Rule::import_desc
                                && c.as_str().trim_start().starts_with("(call_tag")) =>
                    {
                        pre_stmts.push(walk_imported_call_tag(__w, inner)?);
                    }

                    // Active element segment: populate the funcref table so
                    // call_indirect can dispatch through it. Emitted AFTER the
                    // class (a post-stmt) so the `ref.func` tear-off can resolve
                    // each function's chunk — but still before `_start` runs.
                    Rule::elem_field if !is_definition => post_stmts.push(walk_elem_field(__w, inner)?),
                    // A definition's memory/data/table/elem/tag fields fall
                    // through to here: declared, never materialised.
                    _ => {} // type — structural metadata
                }
            }
            _ => {}
        }
    }

    let name = module_name.unwrap_or(default_class_name);
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

    // Register GC struct types (name, parent, field count) so the compiler
    // installs them in the type table with subtype edges — a compile-time
    // directive (`__wast_register_struct_type`) that emits no runtime code.
    // Emitted first so the identity is known before any `struct.new`/`ref.*`.
    let mut result: Vec<Statement> = {
        __w.struct_types
            .iter()
            .map(|(name, parent, fields)| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_struct_type")),
                    args: vec![
                        Argument::positional(Expression::string(name)),
                        Argument::positional(Expression::string(parent.as_deref().unwrap_or(""))),
                        Argument::positional(Expression::int(*fields as i64)),
                    ],
                    optional: false,
                })))
            })
            .collect()
    };
    // Register declared FUNCTION types with their signature, and each defined
    // function with its own, so `ref.test`/`ref.cast` against a concrete
    // `(ref $t)` can match structurally — `Comptype_sub/func` compares the
    // parameter and result TYPES, and no name appears in the rule.
    // Call Tags proposal: emit the `(func … (call_tag $t+))` declarations
    // collected while walking the func fields. They are declarations, so they
    // precede any body that calls through them.
    result.extend(
        std::mem::take(&mut __w.func_call_tag_decls)
            .into_iter()
            .map(|(func, tags)| Statement::new(StmtKind::WasmFuncCallTags { func, tags })),
    );

    result.extend({
        __w.type_func_sigs
            .iter()
            .map(|(name, (params, results))| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_func_type")),
                    args: vec![
                        Argument::positional(Expression::string(name)),
                        Argument::positional(Expression::string(
                            __w.type_func_parent.get(name).map(|s| s.as_str()).unwrap_or(""),
                        )),
                        Argument::positional(Expression::string(&params.join(","))),
                        Argument::positional(Expression::string(&results.join(","))),
                        // The rec group's size and this member's position in
                        // it — see `type_rec_shape`.
                        Argument::positional(Expression::string(
                            __w.type_rec_shape.get(name).map(|s| s.as_str()).unwrap_or(""),
                        )),
                    ],
                    optional: false,
                })))
            })
            .collect::<Vec<_>>()
    });
    // Register GC array types with their element storage type (compile-time
    // directive `__wast_register_array_type(name, elem)`) so the VM can recover
    // the element byte width for `array.init_data`/packed reads. Emitted before
    // the class so the type is element-typed ahead of any `array.*`.
    result.extend({
        __w.array_elem_type
            .iter()
            .map(|(name, elem)| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_array_type")),
                    args: vec![
                        Argument::positional(Expression::string(name)),
                        Argument::positional(Expression::string(elem)),
                    ],
                    optional: false,
                })))
            })
            .collect::<Vec<_>>()
    });
    result.extend(pre_stmts);
    result.push(class);
    // Each defined function's own signature, recorded AFTER the class: the
    // directive resolves the function's method chunk, which does not exist
    // until the class has been compiled.
    let module_class = __w.module_class_name.clone();
    result.extend({
        defined_func_sigs
            .into_iter()
            .map(|(name, (params, results), declared)| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_register_func_sig")),
                    args: vec![
                        Argument::positional(Expression::string(&name)),
                        Argument::positional(Expression::string(&params.join(","))),
                        Argument::positional(Expression::string(&results.join(","))),
                        // ⚠ WHICH MODULE'S CLASS. The compiler resolved the
                        // function against a hardcoded `__wasm_module`, but the
                        // second module's class is `__wasm_module_1` — so every
                        // module after the first registered NO signatures, and
                        // `ref.test (ref $t)` on a function reference there
                        // compared against an empty signature and answered 0.
                        Argument::positional(Expression::string(&module_class)),
                        // The DECLARED type name, for nominal (exact) matching.
                        Argument::positional(Expression::string(
                            declared.as_deref().unwrap_or(""),
                        )),
                    ],
                    optional: false,
                })))
            })
            .collect::<Vec<_>>()
    });
    result.extend(post_stmts);

    // `(start $f)` runs at instantiation: invoke it as a static method of the
    // module class (functions are static methods). This is INDEPENDENT of the
    // `_start` command entry — both run (start first).
    if let Some(sf) = &start_fn_name {
        if start_export_name.as_deref() != Some(sf.as_str()) && !is_definition {
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class_name)),
                    field: sf.clone(),
                    null_safe: false,
                })),
                args: Vec::new(),
                optional: false,
            });
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
            let entry_yields = __w.func_name_results.get(&entry).copied()
                .unwrap_or(0)
                > 0;
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class_name)),
                field: entry,
                null_safe: false,
            });
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args: Vec::new(),
                optional: false,
            });
            let stmt = if entry_yields {
                // `log(entry())` — the entry's declared result is its output.
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("log")),
                    args: vec![Argument::positional(call)],
                    optional: false,
                })
            } else {
                call
            };
            result.push(Statement::new(StmtKind::Expr(stmt)));
        }
    }
    // This module is fully walked: its memories now belong to the script's
    // accumulated index space, so the NEXT module's memidx 0 starts after them.
    // A definition declared none — it must not shift the following module onto
    // a base with nothing allocated under it.
    if !is_definition {
        __w.memory_index_base += module_memory_count;
        __w.table_index_base += module_table_count;
        __w.data_index_base += module_data_count;
        // The elem counter was reset for this module and advanced once per
        // segment while walking it, so it now HOLDS this module's count.
        __w.elem_index_base += __w.elem_seg_counter;
    }
    Ok(result)
}

// ── Function field ────────────────────────────────────────────────────────────

fn walk_func_field(__w: &mut WastWalker, pair: Pair<Rule>, func_index: usize) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut result_count: usize = 0;
    let mut body: Vec<Statement> = Vec::new();
    let mut export_names: Vec<String> = Vec::new();
    let mut labels = LabelStack::new();
    let mut locals_seen: usize = 0;
    let mut local_names: Vec<String> = Vec::new();

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
            // `(func … (call_tag $t+))` — which tags this func's funcref
            // handles. Declaring ANY replaces the default of its own canonical
            // tag, so dropping this on the floor (which is what happened before
            // it was walked) silently left every func universally callable.
            Rule::func_call_tags => {
                let tags: Vec<String> = child
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::index)
                    .map(|c| {
                        let n = c.as_str().trim_start_matches('$').to_string();
                        // An IMPORTED tag's local alias names the exporter's
                        // entity — the same substitution `call_with_tag` does,
                        // and it has to happen here too or the func declares it
                        // handles a different tag than the one callers name.
                        __w.call_tag_alias.get(&n).cloned().unwrap_or(n)
                    })
                    .collect();
                if !tags.is_empty() {
                    // An unnamed `(func (export "e") …)` has no `$id`; its
                    // identity is the export name, which is also what it is
                    // called as a member of the module class. `export_inline`
                    // precedes `func_call_tags` in the grammar, so it is known.
                    let owner = if func_name.is_empty() {
                        export_names.first().cloned().unwrap_or_default()
                    } else {
                        func_name.clone()
                    };
                    if !owner.is_empty() {
                        __w.func_call_tag_decls.push((owner, tags));
                    }
                }
            }
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
                        let pc = __w.type_func_params.get(&qualify_type_name(__w, &sig)).copied()
                            .unwrap_or(0);
                        result_count = __w.type_func_results.get(&qualify_type_name(__w, &sig)).copied()
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
                                is_nullable: false,
                            })
                            .collect();
                    }
                }
            }
            Rule::local => {
                let (decls, names) = walk_local(child, params.len() + locals_seen)?;
                locals_seen += decls.len();
                local_names.extend(names);
                body.extend(decls);
            }
            Rule::instr => {
                instr_pairs.push(child);
            }
            _ => {}
        }
    }

    // Expose the function's result count so a `return` inside a multi-value
    // function reraises the top N values as a uniform tuple (multi-value ABI).
    __w.current_fn_results = result_count;
    // Params first, then locals — one index space, and the ONLY thing that can
    // answer `local.get 0` when the param is named.
    __w.local_index_name = params
        .iter()
        .map(|p| p.name.clone())
        .chain(local_names.iter().cloned())
        .collect();
    body.extend(fold_instructions(__w, instr_pairs, &mut labels)?);
    __w.local_index_name.clear();
    __w.current_fn_results = 0;

    if func_name.is_empty() {
        // Same rule the prescan's index→name map used: id, else the first
        // inline export NO OTHER FUNCTION CLAIMS AS ITS `$id`, else a
        // per-INDEX synthetic name. A module may hold several unnamed
        // functions, so a shared constant would collide (and `(export "e"
        // (func N))` could not tell them apart); and an export name that is
        // also someone's id would make the two functions ONE — see the
        // namespace note in the prescan.
        func_name = export_names
            .iter()
            .find(|e| !__w.defined_func_names.contains(*e))
            .cloned()
            .unwrap_or_else(|| format!("__wasm_func_{func_index}"));
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
            is_sub: false,
        },
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
            is_nullable: false,
        }]);
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
            is_nullable: false,
        })
        .collect())
}

fn walk_local(pair: Pair<Rule>, index_base: usize) -> Result<(Vec<Statement>, Vec<String>), String> {
    let mut name: Option<String> = None;
    let mut types: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::any_val_type | Rule::val_type => types.push(child.as_str().to_string()),
            _ => {}
        }
    }
    let mut minted: Vec<String> = Vec::with_capacity(types.len());
    let decls: Vec<Statement> = types
        .iter()
        .enumerate()
        .map(|(i, t)| {
            // An anonymous local is addressed by INDEX, in the same index
            // space as the params, so its name is `p<params + locals-so-far>`
            // — never a per-group count. A NAMED one keeps its `$id`, and the
            // index spelling reaches it through `local_index_name`.
            let var_name = name
                .clone()
                .unwrap_or_else(|| format!("p{}", index_base + i));
            minted.push(var_name.clone());
            let init = match t.as_str() {
                "i32" => Expression::int(0),
                // An i64 zero is a BigInt — the exact-integer shape every
                // other i64 value carries (a plain Int literal compiles to
                // f64 and compares unequal to `(i64.const 0)`).
                "i64" => Expression::with_span(ExprKind::Lit(Literal::BigInt(0)), Span::default()),
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
                    optional: false,
                }),
                _ => Expression::null(),
            };
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(var_name),
                    type_hint: Some(t.clone().into()),
                    init: Some(init),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            })
        })
        .collect();
    Ok((decls, minted))
}

// ── Instructions ──────────────────────────────────────────────────────────────
//
// WAT instructions that produce a control-flow effect (block, loop, br, br_if,
// return, unreachable) are lowered to proper AST *statements* so the compiler
// emits the right WASM structured-control opcodes (BLOCK, LOOP, BR, BR_IF,
// RETURN, UNREACHABLE).  Value-producing instructions become expressions.

#[allow(dead_code)]
fn walk_instr_as_stmts(__w: &mut WastWalker, 
    pair: Pair<Rule>,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_stmts(__w, inner, span, labels),
        Rule::plain_instr => walk_plain_instr_as_stmts(__w, inner, span, labels),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
    }
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
        optional: false,
    })
}

fn walk_instr_as_expr(__w: &mut WastWalker, pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("Empty instr")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(__w, inner, span, labels),
        Rule::plain_instr => walk_plain_instr_as_expr(__w, inner, span, labels),
        _ => Err(format!("Unexpected instr rule: {:?}", inner.as_rule())),
    }
}

// ── Plain instructions ────────────────────────────────────────────────────────

#[allow(dead_code)]
fn walk_plain_instr_as_stmts(__w: &mut WastWalker, 
    pair: Pair<Rule>,
    _span: Span,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    fold_instructions(__w, vec![pair], labels)
}

fn walk_plain_instr_as_expr(__w: &mut WastWalker, 
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
    let name = peel_mem_selector(__w, &name, &mut raw_args, labels)?;
    let mut args: Vec<Expression> = Vec::new();
    for raw in raw_args {
        args.push(walk_instr_arg_for(__w, raw, labels, &name)?);
    }
    map_instr_to_ast(__w, name, args, span)
}

// ── Folded instructions ───────────────────────────────────────────────────────

#[allow(dead_code)]
fn walk_folded_instr_as_stmts(__w: &mut WastWalker, 
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
            let effective = labels.push(__w, label.clone(), LabelKind::Block, Vec::new());
            let body = fold_instructions(__w, instr_pairs, labels)?;
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
            let effective = labels.push(__w, label.clone(), LabelKind::Loop, Vec::new());
            let mut body = fold_instructions(__w, instr_pairs, labels)?;
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
                .map(|c| walk_instr_as_expr(__w, c, labels))
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
                        cond = Some(walk_instr_as_expr(__w, child, labels)?);
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
            let expr = walk_folded_core(__w, name, all_children, span, labels)?;
            Ok(vec![Statement::with_span(StmtKind::Expr(expr), span)])
        }
    }
}

fn walk_folded_instr_as_expr(__w: &mut WastWalker, 
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
    walk_folded_core(__w, name, all_children, span, labels)
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
fn walk_folded_core(__w: &mut WastWalker,
    name: String,
    children: Vec<Pair<Rule>>,
    span: Span,
    labels: &mut LabelStack,
) -> Result<Expression, String> {
    // ── Multi-memory selector ─────────────────────────────────────────────
    // A FOLDED memory op carries its memidx the same way the plain form does:
    // as leading BARE `instr_arg` immediates, ahead of the folded operands
    // (`(i32.store 1 (i32.const 0) (i32.const 42))`). Peel them into the
    // `@@mem<N>` name suffix the emitter reads. Without this the selector
    // reached the emitter as an OPERAND and every folded access ran against
    // memory 0.
    let (name, children) = if mem_op_immediate_count(&name) > 0 {
        let mut children = children;
        let mut selectors: Vec<Pair<Rule>> = Vec::new();
        // Collect the memidx immediates AND any bare index immediate that
        // follows them (`memory.init`'s dataidx), because deciding whether a
        // leading index IS a memidx needs to see how many there are in total.
        // `peel_mem_selector` consumes only the ones it claims.
        let want = mem_op_immediate_count(&name) + mem_op_trailing_index_count(&name);
        while selectors.len() < want
            && children
                .first()
                .map(|c| c.as_rule() == Rule::instr_arg && is_bare_index_arg(c))
                .unwrap_or(false)
        {
            selectors.push(children.remove(0));
        }
        let name = peel_mem_selector(__w, &name, &mut selectors, labels)?;
        // Whatever it declined to consume is a real operand — put it back in
        // front, in order.
        for (at, left) in selectors.into_iter().enumerate() {
            children.insert(at, left);
        }
        (name, children)
    } else {
        (name, children)
    };
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
        labels.push(__w, label.clone(), kind.clone(), Vec::new());
        let body = fold_instructions(__w, instr_pairs, labels)?;
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
            Rule::instr_arg => args.push(walk_instr_arg_for(__w, child, labels, &name)?),
            Rule::instr => args.push(walk_instr_as_expr(__w, child, labels)?),
            Rule::then_block => {
                has_then = true;
                let mut instr_pairs = Vec::new();
                for sub in child.into_inner() {
                    if sub.as_rule() == Rule::instr {
                        instr_pairs.push(sub);
                    }
                }
                let body = fold_instructions(__w, instr_pairs, labels)?;
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
                let body = fold_instructions(__w, instr_pairs, labels)?;
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
                else_: Box::new(else_val),
            },
            span,
        ));
    }

    map_instr_to_ast(__w, name, args, span)
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

/// Strip the `offset=`/`align=` memarg string-args off a load/store's operand
/// list, returning the remaining operands and the `offset=` value.
///
/// ⛔ The offset is NOT folded into the address. WASM computes the effective
/// address in UNBOUNDED arithmetic over the address read as UNSIGNED
/// (§4.4.7): `i32.load offset=25` at address `-1` is `4294967295 + 25`, which
/// is out of bounds and TRAPS. A folded AST `addr + 25` is a SIGNED add, so it
/// computed 24, stayed in bounds and returned a byte — `address.wast`'s whole
/// `-1` block. The offset therefore rides in the instruction's MEMARG, where
/// the VM's `effective_addr` already does the unsigned widen and a saturating
/// add. It is carried across as an `@@off<N>` name suffix, the same channel
/// `@@mem<N>` uses, because an extra AST argument would shift every
/// immediate-position index in the emitter's `OperandFormat` match.
///
/// Folding was also blind in the PLAIN spelling, where the address is not an
/// argument at all — it is on the enclosing block's stack — so there was no
/// slot to fold into and the offset was silently dropped.
///
/// `align=` is a pure hint (WASM validates it but the semantics do not depend
/// on it) and is dropped.
fn strip_memarg(args: Vec<Expression>) -> (Vec<Expression>, u64) {
    let mut offset: u64 = 0;
    let mut rest: Vec<Expression> = Vec::new();
    for a in args {
        if let ExprKind::Lit(Literal::Str(s)) = &a.kind {
            if let Some(n) = s.strip_prefix("offset=") {
                offset = offset.saturating_add(parse_memarg_number(n));
                continue;
            }
            if s.starts_with("align=") {
                continue;
            }
        }
        rest.push(a);
    }
    (rest, offset)
}

/// A memarg field is a WAT `uN`: decimal or `0x`-hex, either with `_` digit
/// separators. `parse::<u64>()` alone read `offset=0x008` as 0 — silently, and
/// the whole of `align.wast` writes its offsets in hex.
fn parse_memarg_number(text: &str) -> u64 {
    let t = text.replace('_', "");
    match t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
        Some(hex) => u64::from_str_radix(hex, 16).unwrap_or(0),
        None => t.parse::<u64>().unwrap_or(0),
    }
}

/// The binding name a numeric `local.get`/`global.get` (and set) index lowers
/// to. Locals/params use the `p<i>` name space; globals resolve through the
/// declaration-order `GLOBAL_INDEX_NAME` (falling back to the same synthetic
/// scheme the pre-scan used, so it works even if the pre-scan missed one).
fn index_binding_name(__w: &mut WastWalker, i: i64, is_global: bool) -> String {
    if is_global {
        __w.global_index_name.get(i as usize).cloned()
            .unwrap_or_else(|| format!("__wasm_global_{i}"))
    } else {
        local_index_binding_name(__w, i)
    }
}

/// The binding a numeric local/param index names. WAT lets ANY local be
/// addressed by index — `(param $a i32)` is still local 0 — so the declared
/// spelling has to be looked up. `p<i>` is the fallback for an index the
/// declaration walk never saw (and is what an UNNAMED local is called anyway).
fn local_index_binding_name(__w: &WastWalker, i: i64) -> String {
    usize::try_from(i)
        .ok()
        .and_then(|i| __w.local_index_name.get(i).cloned())
        .unwrap_or_else(|| format!("p{i}"))
}

/// Resolve a memory-index immediate: a literal integer, or a `$name` looked up
/// in the declaration-order `MEMORY_NAME_INDEX`. Anything else is memidx 0.
///
/// A numeric memidx is MODULE-RELATIVE, and `memory_slots` is what it is
/// relative to — a plain `base + n` was only ever right while every memory was
/// freshly declared, because an imported one occupies an index without
/// allocating a slot. Named memories were registered already resolved.
fn resolve_wat_memidx(__w: &mut WastWalker, e: &Expression) -> usize {
    let default = default_memory_slot(__w);
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => __w
            .memory_slots
            .get(*n as usize)
            .copied()
            .unwrap_or(__w.memory_index_base + *n as usize),
        ExprKind::Ident(nm) => __w.memory_name_index.get(nm).copied().unwrap_or(default),
        _ => default,
    }
}

/// The script slot this module's memidx 0 names — its own first memory, which
/// may be one it IMPORTED rather than one it declared.
fn default_memory_slot(__w: &WastWalker) -> usize {
    __w.memory_slots
        .first()
        .copied()
        .unwrap_or(__w.memory_index_base)
}

/// Number of leading memory-index immediates a memory op may carry: `memory.copy`
/// names two memories (dst, src); every other memory op names at most one. 0 =
/// not a memory op (never peels a selector).
fn mem_op_immediate_count(name: &str) -> usize {
    match name {
        "memory.copy" => 2,
        "memory.init"
        | "memory.fill" | "memory.size" | "memory.grow" | "i32.load" | "i64.load" | "f32.load"
        | "f64.load" | "i32.load8_s" | "i32.load8_u" | "i32.load16_s" | "i32.load16_u"
        | "i64.load8_s" | "i64.load8_u" | "i64.load16_s" | "i64.load16_u" | "i64.load32_s"
        | "i64.load32_u" | "i32.store" | "i64.store" | "f32.store" | "f64.store" | "i32.store8"
        | "i32.store16" | "i64.store8" | "i64.store16" | "i64.store32" => 1,
        _ => 0,
    }
}

/// Which argument slot of an op holds a DATA SEGMENT index, if any. The
/// `@@mem` suffix is already on the name by the time this is asked, so match on
/// the head rather than the whole string.
fn dataidx_arg_position(name: &str) -> Option<usize> {
    let base = name.split_once("@@").map(|(b, _)| b).unwrap_or(name);
    match base {
        "memory.init" | "data.drop" => Some(0),
        // `array.new_data $T $d` / `array.init_data $T $d` — typeidx first.
        "array.new_data" | "array.init_data" => Some(1),
        _ => None,
    }
}

/// Resolve a written elemidx — `$e` (named) or `2` (numeric) — into the
/// SCRIPT's element-segment index space, exactly as `resolve_wat_dataidx` does
/// for data segments. Before this, a named segment reached the emitter as an
/// `Ident` and encoded as 0, and a numeric one was never shifted.
fn resolve_wat_elemidx(__w: &mut WastWalker, e: &Expression) -> i64 {
    let base = __w.elem_index_base as i64;
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => base + *n,
        ExprKind::Ident(nm) => __w
            .elem_name_index
            .get(nm)
            .copied()
            .map(|i| i as i64)
            .unwrap_or(base),
        _ => base,
    }
}

/// Which argument slot of an op holds an ELEMENT SEGMENT index, if any.
/// `table.init` is not here: its elemidx position depends on whether a tableidx
/// was written, so it is resolved in its own arm below.
fn elemidx_arg_position(name: &str) -> Option<usize> {
    match name {
        "elem.drop" => Some(0),
        // `array.new_elem $T $e` / `array.init_elem $T $e` — typeidx first.
        "array.new_elem" | "array.init_elem" => Some(1),
        _ => None,
    }
}

/// Resolve a written dataidx — `$d` (named) or `2` (numeric) — into the
/// SCRIPT's data-segment index space. A numeric index is module-relative and
/// shifts by `DATA_INDEX_BASE`; a named segment was registered pre-shifted.
fn resolve_wat_dataidx(__w: &mut WastWalker, e: &Expression) -> i64 {
    let base = __w.data_index_base as i64;
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => base + *n,
        ExprKind::Ident(nm) => __w
            .data_name_index
            .get(nm)
            .copied()
            .map(|i| i as i64)
            .unwrap_or(base),
        _ => base,
    }
}

/// Bare index immediates that FOLLOW the memidx. Only `memory.init` has one:
/// its dataidx. It is what makes the memidx AMBIGUOUS — `memory.init $d` names
/// a data segment on the default memory, `memory.init $m $d` names both — and
/// the only thing that tells them apart is HOW MANY bare indices are written.
/// Peeling unconditionally read `$d` as a memory and left the op with no data
/// segment; not peeling at all sent every `memory.init $mem2 $d` to memory 0,
/// which is `multi-memory/memory-multi.wast`.
fn mem_op_trailing_index_count(name: &str) -> usize {
    if name == "memory.init" { 1 } else { 0 }
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
fn peel_mem_selector(__w: &mut WastWalker, 
    name: &str,
    raw_args: &mut Vec<Pair<Rule>>,
    labels: &mut LabelStack,
) -> Result<String, String> {
    let n = mem_op_immediate_count(name);
    if n == 0 {
        return Ok(name.to_string());
    }
    // An op whose memidx is followed by another bare index (`memory.init`'s
    // dataidx) only HAS a memidx when both are written. Decline without
    // consuming anything, so the caller's put-back leaves the dataidx where the
    // emitter expects it — as the first arg.
    let trailing = mem_op_trailing_index_count(name);
    if trailing > 0 {
        let bare = raw_args
            .iter()
            .take_while(|r| is_bare_index_arg(r))
            .count();
        if bare < n + trailing {
            return Ok(name.to_string());
        }
    }
    let mut indices = Vec::new();
    while indices.len() < n && raw_args.first().map(is_bare_index_arg).unwrap_or(false) {
        let r = raw_args.remove(0);
        let e = walk_instr_arg_pair(__w, r, labels)?;
        indices.push(resolve_wat_memidx(__w, &e));
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

/// A table op's immediate/operand shape: `(max table-index immediates, stack
/// operands)`. Every table op carries its tableidx as a fixed U16 IMMEDIATE
/// (`core_ops`/`misc`: `table.get`/`set`/`size`/`grow`/`fill` = one, `table.copy`
/// = two, dst then src — the order `Op::TABLE_COPY` reads them in). `table.init`
/// and `elem.drop` are NOT here: their leading immediate is an elemidx, and
/// `table.init` has its own [elem, table] normalization below.
fn table_op_shape(name: &str) -> Option<(usize, usize)> {
    match name {
        "table.size" => Some((1, 0)),
        "table.get" => Some((1, 1)),
        "table.set" | "table.grow" => Some((1, 2)),
        "table.fill" => Some((1, 3)),
        "table.copy" => Some((2, 3)),
        _ => None,
    }
}

/// Resolve a written tableidx immediate — `$t3` (named) or `2` (numeric) — into
/// the SCRIPT's table index space. A numeric index is module-relative and shifts
/// by `TABLE_INDEX_BASE`; a named table was registered pre-shifted.
fn resolve_wat_tableidx(__w: &mut WastWalker, e: &Expression) -> i64 {
    let base = __w.table_index_base as i64;
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => base + *n,
        ExprKind::Ident(nm) => resolve_table_index(__w, nm),
        _ => base,
    }
}

/// The operand of `f32.const` / `f64.const` as a FLOAT, even when it was
/// written in integer form — `(f64.const 1)` is the double 1.0, not the
/// integer 1, and the spec's suite writes plenty of them that way.
///
/// `instr_arg_inner_to_expr` walks operands without knowing which instruction
/// they belong to, so an integer-form float constant arrives here as
/// `Lit(Int)`. Left alone it stayed an integer all the way down and the
/// conversion instructions then read it at the wrong type.
///
/// Magnitudes at or above 2^64 arrive as `Lit(Float)` already: they cannot be
/// an integer const of any width, so `parse_integer` reads them as floats
/// rather than saturating to 0 (that is `float_literals.wast:143`,
/// `(f32.const 0x1_0000_0000_0000_0000_0000)` = 2^80).
///
/// KNOWN LIMIT: the one remaining ambiguous window is [2^63, 2^64) in a MODULE
/// BODY. `(i64.const 18446744073709551615)` legally means -1, so that range
/// must keep wrapping — which makes a wrapped value indistinguishable from a
/// genuinely negative one without the source text, and this layer does not
/// receive it. The expectation / `invoke`-argument side does have the text and
/// handles the range correctly (see the `f32.const`/`f64.const` case in the
/// const-expr walker), and that is where the spec's own `conversions.wast:246`
/// sits. Closing it here needs the instruction name plumbed into the operand
/// walker, so that a float const can opt out of the wrap.
fn float_const_operand(args: Vec<Expression>, span: Span) -> Expression {
    match args.into_iter().next() {
        Some(e) => match &e.kind {
            ExprKind::Lit(Literal::Int(n)) => {
                Expression::with_span(ExprKind::Lit(Literal::Float(*n as f64)), span)
            }
            _ => e,
        },
        None => Expression::float(0.0),
    }
}

/// How an integer-const rejection names what it was actually handed.
///
/// The message has to quote the FORM, because the two cases that reach it look
/// nothing alike in source: `nan:arithmetic` (a float NaN spelling) and `1.5`.
fn describe_const_operand(v: &Expression) -> String {
    match &v.kind {
        ExprKind::Lit(Literal::Float(f)) if f.is_nan() => "a NaN".to_string(),
        ExprKind::Lit(Literal::Float(f)) => format!("the float {f}"),
        ExprKind::Lit(Literal::Str(s)) => format!("the string {s:?}"),
        _ => "a non-integer operand".to_string(),
    }
}

/// `call_ref $sig` / `return_call_ref $sig` → the WASM instruction's own
/// opcode, with the `[funcref, operand*]` stack layout `call_value` expects —
/// the same shape `emit_call` has always emitted.
///
/// ⛔ These used to lower to a generic `ExprKind::Call` whose callee was an
/// EXPRESSION, which sends the compiler down its dynamic-callee ladder. That
/// ladder opens by reading `__vybe_method_receiver` OFF the callee — a
/// `struct.get` — so `call_ref` on a null funcref trapped "null structure
/// reference (struct.get)" before anything reached a call path at all
/// (`call_ref.wast`). The message was honest: it was not a call. The tail form
/// went through `__wasm_return_call`, which wants a QUALIFIED callee name and
/// got a value, and reported "null is not callable"
/// (`return_call_ref.wast`).
///
/// ⛔ Both were also blind in the PLAIN spelling, where the operands are on the
/// enclosing block's stack and never appear as arguments: popping the funcref
/// off an empty argument list produced a call to null. There the argument count
/// has to come from `$sig`, which is the only thing that knows it.
fn build_call_ref_opcode(
    __w: &mut WastWalker,
    name: &str,
    args: Vec<Expression>,
    span: Span,
) -> Result<Expression, String> {
    // ⛔ `resolve_wast_type_name` is THE answer to "which type is `$sig`" — it
    // takes the NAMED and the POSITIONAL spelling both (`call_ref 0`, which the
    // spec writes wherever it writes numeric type indices). Matching only
    // `ExprKind::Ident` here would have been a fourth private answer to a
    // question this already answers, and it would have missed silently: a
    // numeric `$sig` would resolve to 0 params.
    let type_name = resolve_wast_type_name(__w, args.first());
    let sig_params = __w.type_func_params.get(&type_name).copied();
    let sig_results = __w.type_func_results.get(&type_name).copied();

    let mut rest = args;
    if !rest.is_empty() {
        rest.remove(0); // the `$sig` type immediate is not a stack value
    }
    // FOLDED: `[operand*, funcref]` — the funcref is written LAST and the
    // opcode wants it FIRST, below its arguments. Its count is what was
    // actually written, not what `$sig` declares, so an unresolved `$sig`
    // cannot silently change the folded call's arity.
    // FLAT: nothing was folded in, the operands are already on the stack in
    // the right order, and only the count is needed.
    let callee = rest.pop();
    let argc = match (&callee, sig_params) {
        (Some(_), _) => rest.len(),
        (None, Some(n)) => n,
        // ⛔ NOT a silent 0. Nothing was folded in AND `$sig` did not resolve,
        // so neither source of the arity is available and any number here is a
        // guess that would mis-slice the stack at run time.
        (None, None) => {
            return Err(format!(
                "{name}: cannot determine the argument count — `{type_name}` is not a \
                 declared function type and no operands were folded in"
            ))
        }
    };
    // The result count rides along for the WASM writer, which picks the
    // call_indirect functype from the `(argc, results)` pair — so a wrong value
    // here writes a wrong binary even though the VM ignores it (the callee
    // chunk's own arity drives execution). Default to the declared type's
    // result count; fall back to 1 only when `$sig` gave us nothing, which the
    // arity check above has already made unreachable for the flat spelling.
    let sig_results = sig_results.unwrap_or(1);
    let mut operands = vec![
        Expression::int(argc as i64),
        Expression::int(sig_results as i64),
    ];
    if let Some(callee) = callee {
        operands.push(callee);
        operands.extend(rest);
    }
    Ok(make_call(name, operands, span))
}

fn map_instr_to_ast(__w: &mut WastWalker, name: String, args: Vec<Expression>, span: Span) -> Result<Expression, String> {
    // A memory op with NO explicit selector targets ITS module's memory 0 —
    // which, in a multi-module script, is program memory <base>. Explicit
    // selectors were already `@@mem`-mangled (shifted) at the parse sites.
    // ⚠ THE PLAIN AND FOLDED SPELLINGS MUST NAME THE SAME TYPE.
    //
    // `(ref.cast (ref $T) …)` reaches the walker through `ref_type_arg`, which
    // qualifies. The unfolded `ref.cast $T` carries its type as a bare `id`
    // instr_arg instead, so it arrived UNQUALIFIED and looked up a name the
    // registration no longer uses — `trap: ref.cast failed: value is not Sub`,
    // with the bare `Sub` in the message as the tell.
    //
    // Only the leading TYPE operand is rewritten, and only when it is an
    // `Ident` (a `$id`): abstract spellings arrive as strings, and
    // `qualify_type_name` is idempotent, so a folded arg that came in already
    // qualified passes through untouched.
    let args = if matches!(
        name.as_str(),
        "ref.test" | "ref.test_null" | "ref.cast" | "ref.cast_null" | "ref.cast_desc_eq"
            // The ARRAY family all lead with a typeidx. Several of their arms
            // pass that argument through VERBATIM rather than via
            // `resolve_wast_type_name` — `array.new_default` hands it straight
            // to `array_new` — so the registration was qualified while the
            // reference stayed bare and the element width was looked up under a
            // name that no longer existed. Wrong width, wrong bounds:
            // `array.init_data: source out of bounds` on an array that fits.
            | "array.new"
            | "array.new_default"
            | "array.new_fixed"
            | "array.new_data"
            | "array.new_elem"
            | "array.init_data"
            | "array.init_elem"
            | "array.get"
            | "array.get_s"
            | "array.get_u"
            | "array.set"
            | "array.fill"
    ) {
        let mut a = args;
        if let Some(first) = a.first_mut() {
            if let ExprKind::Ident(n) = &first.kind {
                let q = qualify_type_name(__w, n);
                *first = Expression::ident(&q);
            }
        }
        a
    } else {
        args
    };
    let default_mem = default_memory_slot(__w);
    let name = if default_mem > 0 && !name.contains("@@mem") && mem_op_immediate_count(&name) > 0 {
        if name == "memory.copy" {
            format!("memory.copy@@mem{default_mem}@@mem{default_mem}")
        } else {
            format!("{name}@@mem{default_mem}")
        }
    } else {
        name
    };
    // A constant `offset=N` memarg on a load/store rides across as an
    // `@@off<N>` name suffix and is emitted IN the instruction's memarg — see
    // `strip_memarg` for why folding it into the address was wrong. The
    // suffixes only ever APPEND, and every reader takes the base name as
    // everything before the first `@@`, so `@@off` and `@@mem` compose in
    // either order.
    let (name, args) = if name.contains(".load") || name.contains(".store") {
        let (rest, offset) = strip_memarg(args);
        let name = if offset != 0 {
            format!("{name}@@off{offset}")
        } else {
            name
        };
        (name, rest)
    } else {
        (name, args)
    };
    // A DATAIDX immediate is module-relative over a script-wide segment list,
    // the same split the memidx/tableidx bases exist for. `memory.init` and
    // `data.drop` name it first; the GC array-from-data ops name a typeidx
    // first and the dataidx second. A written `$d` resolves through the
    // pre-scan, a written number shifts by the base.
    let args = match dataidx_arg_position(&name) {
        Some(at) => {
            let mut args = args;
            if let Some(written) = args.get(at).cloned() {
                args[at] = Expression::with_span(
                    ExprKind::Lit(Literal::Int(resolve_wat_dataidx(__w, &written))),
                    span,
                );
            }
            args
        }
        None => args,
    };
    // …and the same for an ELEMIDX immediate.
    let args = match elemidx_arg_position(&name) {
        Some(at) => {
            let mut args = args;
            if let Some(written) = args.get(at).cloned() {
                args[at] = Expression::with_span(
                    ExprKind::Lit(Literal::Int(resolve_wat_elemidx(__w, &written))),
                    span,
                );
            }
            args
        }
        None => args,
    };
    // NOTE: multi-memory selectors (`i32.store 1`, `memory.copy 1 0`, …) are
    // peeled off at the plain-instruction parse site (`peel_mem_selector`), where
    // a genuine memidx immediate — a BARE `integer`/`id` token — is distinguishable
    // from a folded operand the WAT grammar greedily attaches as an `instr_arg`.
    // By the time args reach here they are already `name@@mem<N>`-mangled with the
    // memidx stripped, so no arity-based inference (which cannot tell an immediate
    // from a flushed operand) is done in this shared lowering.
    match name.as_str() {
        // `table.get/set/size/grow/fill/copy` — the tableidx is an OPTIONAL
        // leading immediate in WAT but a MANDATORY U16 in the bytecode, so
        // normalize to exactly `max_imm` resolved integers followed by the
        // stack operands. Without this a written `$t3` reached the emitter as an
        // `Ident` (encoded as immediate 0 → every op hit table 0), and an
        // OMITTED index let the first stack operand be read as the immediate and
        // then never pushed. table_get.wast:9 is both at once.
        n if table_op_shape(n).is_some() => {
            let (max_imm, stack_ops) = table_op_shape(&name).expect("checked by the guard");
            let mut a = args;
            // Count, capped: an over-full arg list cannot invent a selector.
            // An `Ident` head is NOT a usable "this is a selector" signal —
            // `local.get $i` lowers to `Ident` too, so keying on it made
            // `(table.get (local.get $i))` read the LOCAL as the tableidx.
            let n_idx = a.len().saturating_sub(stack_ops).min(max_imm);
            let mut idx: Vec<Expression> = a
                .drain(..n_idx)
                .map(|e| Expression::int(resolve_wat_tableidx(__w, &e)))
                .collect();
            // An OMITTED selector means THIS module's table 0 — which, in a
            // multi-module script, is program table <base>.
            let default_idx = __w.table_index_base as i64;
            idx.resize(max_imm, Expression::int(default_idx));
            idx.append(&mut a);
            return Ok(make_call(&name.replace('.', "_"), idx, span));
        }
        // `table.init tableidx? elemidx` — WAT allows 1 index (elemidx, table 0)
        // or 2 (tableidx elemidx). The VM's TABLE_INIT reads two u16 immediates
        // in [elem_idx, table_idx] order, then 3 stack operands (dst, src, len).
        // Normalize the leading immediates to exactly [elem, table] so the table
        // index is never mistaken for the first stack operand.
        "table.init" => {
            let mut a = args;
            let n_idx = a.len().saturating_sub(3); // 3 stack operands
            // The DEFAULT table is this module's table 0 = program table <base>;
            // an explicit tableidx resolves through the same script-wide space.
            let base = Expression::int(__w.table_index_base as i64);
            let (elem, table) = match n_idx {
                0 => (Expression::int(__w.elem_index_base as i64), base),
                1 => {
                    let written = a.remove(0);
                    (
                        Expression::int(resolve_wat_elemidx(__w, &written)),
                        base,
                    )
                }
                _ => {
                    let table = a.remove(0); // text order: tableidx first
                    let written = a.remove(0); // then elemidx
                    (
                        Expression::int(resolve_wat_elemidx(__w, &written)),
                        Expression::int(resolve_wat_tableidx(__w, &table)),
                    )
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
                ExprKind::Ident(n) => __w.array_elem_type.get(&qualify_type_name(__w, n)).cloned(),
                _ => None,
            });
            let rest: Vec<Expression> = args.into_iter().skip(1).collect();
            let get = make_call("array_get", rest, span);
            Ok(match (elem.as_deref(), signed) {
                (Some("i8"), true) => make_call("i32.extend8_s", vec![get], span),
                (Some("i16"), true) => make_call("i32.extend16_s", vec![get], span),
                (Some("i8"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(get),
                    right: Box::new(Expression::int(0xFF)),
                }),
                (Some("i16"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(get),
                    right: Box::new(Expression::int(0xFFFF)),
                }),
                // Non-packed (i32/i64/ref/…): the stored value is already exact.
                _ => get,
            })
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
            match &v.kind {
                ExprKind::Lit(Literal::Int(n)) => {
                    let reinterp = (*n as u32) as i32 as i64;
                    Ok(Expression::with_span(
                        ExprKind::Lit(Literal::Int(reinterp)),
                        span,
                    ))
                }
                ExprKind::Lit(Literal::BigInt(n)) => Ok(Expression::with_span(
                    ExprKind::Lit(Literal::Int((*n as u32) as i32 as i64)),
                    span,
                )),
                _ => Err(format!(
                    "i32.const takes an integer literal, not {}",
                    describe_const_operand(&v)
                )),
            }
        }
        // i64.const needs an exact 64-bit value; a plain Int literal compiles to
        // f64 (losing bits past 2^53), so carry it as the exact-integer literal
        // the i64 opcodes read via `as_i64`.
        "i64.const" => {
            let v = args.into_iter().next().unwrap_or(Expression::int(0));
            match &v.kind {
                ExprKind::Lit(Literal::Int(n)) => Ok(Expression::with_span(
                    ExprKind::Lit(Literal::BigInt(*n)),
                    span,
                )),
                ExprKind::Lit(Literal::BigInt(_)) => Ok(v),
                _ => Err(format!(
                    "i64.const takes an integer literal, not {}",
                    describe_const_operand(&v)
                )),
            }
        }
        "f64.const" => Ok(float_const_operand(args, span)),
        // f32.const carries an f32 value: demote the (exact-text) f64 literal to
        // single precision so it lands as `Value::F32`, matching WASM. NaNs take
        // the bit-exact route instead — see `f32_const_expr`.
        "f32.const" => Ok(f32_const_expr(float_const_operand(args, span), span)),
        // wasm:js-string builtins — string.const "text" → string literal
        "string.const" => Ok(args.into_iter().next().unwrap_or(Expression::string(""))),

        // ── Local / global get → Ident ────────────────────────────────────
        // A numeric index names a LOCAL/param (`p<i>`) for `local.get`, but a
        // GLOBAL by declaration index for `global.get` (separate name spaces).
        "local.get" | "global.get" => {
            let is_global = name == "global.get";
            let idx = args.into_iter().next().unwrap_or(Expression::int(0));
            Ok(match &idx.kind {
                // An imported global is a second NAME for the exporter's cell.
                ExprKind::Ident(n) if is_global && __w.global_import_alias.contains_key(n) => {
                    let target = __w.global_import_alias[n].clone();
                    Expression::with_span(ExprKind::Ident(target), span)
                }
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(index_binding_name(__w, *i, is_global)), span)
                }
                _ => idx,
            })
        }

        // ── Local / global set → Assign ───────────────────────────────────
        "local.set" | "global.set" => {
            let is_global = name == "global.set";
            let mut it = args.into_iter();
            let target_raw = it.next().unwrap_or(Expression::int(0));
            let value = it.next().unwrap_or(Expression::null());
            let target = match &target_raw.kind {
                // Writing an imported mutable global writes the exporter's
                // cell — the spec shares one cell, it does not copy.
                ExprKind::Ident(n) if is_global && __w.global_import_alias.contains_key(n) => {
                    let t = __w.global_import_alias[n].clone();
                    Expression::with_span(ExprKind::Ident(t), span)
                }
                ExprKind::Ident(n) => Expression::with_span(ExprKind::Ident(n.clone()), span),
                ExprKind::Lit(Literal::Int(i)) => {
                    Expression::with_span(ExprKind::Ident(index_binding_name(__w, *i, is_global)), span)
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
                // Same index space, same lookup — a `local.tee 0` into a NAMED
                // param wrote to a binding called `p0` that nothing reads.
                ExprKind::Lit(Literal::Int(i)) => local_index_binding_name(__w, *i),
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
                    else_: Box::new(val2),
                },
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
            // `call <funcidx>` — a NUMERIC index, which is what the spec's own
            // suite writes (`fac.wast`'s `(call 0)` recursing into itself). The
            // index names a function; without resolving it the literal became
            // the callee and every such call died with "f64 is not callable".
            //
            // `FUNC_INDEX_NAME` is already the spec's function index space —
            // imports pushed first, then defined functions, in module order —
            // so this is a lookup, not a second table. Out of range stays a
            // literal: `call` past the index space is a validation error, and
            // inventing a name here would hide it behind a wrong callee.
            let callee = match &callee.kind {
                ExprKind::Lit(Literal::Int(idx)) if *idx >= 0 => __w.func_index_name.get(*idx as usize).cloned()
                    .map_or(callee, |n| Expression::with_span(ExprKind::Ident(n), span)),
                _ => callee,
            };
            // A call to a function DEFINED in this module targets a static method
            // of the module class; qualify `Ident(f)` as `ClassName.f`. Imports
            // keep their bare name so the profile builtin table resolves them.
            let callee = match &callee.kind {
                // An IMPORT is a second name for an EXPORTING module's method,
                // so it must be qualified against THAT class — checked first,
                // because the importing module also declares a same-named stub
                // that `DEFINED_FUNC_NAMES` would otherwise claim.
                ExprKind::Ident(n)
                    if __w.import_alias.contains_key(n) =>
                {
                    let (class, method) =
                        __w.import_alias.get(n).cloned().expect("just checked");
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(Expression::ident(&class)),
                            field: method,
                            null_safe: false,
                        },
                        span,
                    )
                }
                // A host import: lower to the `host:<module>:<fn>` callee built
                // from its own declaration. Checked BEFORE `defined_func_names`
                // for the same reason the registered-module arm is — the
                // importing module also declares a same-named stub that would
                // otherwise claim it. Ranked after that arm because an import
                // satisfied by another module in the same script is a real
                // cross-module call, not a host call.
                ExprKind::Ident(n) if __w.host_import_alias.contains_key(n) => {
                    let target = __w.host_import_alias.get(n).cloned().expect("just checked");
                    Expression::with_span(ExprKind::Ident(target), span)
                }
                ExprKind::Ident(n) if __w.defined_func_names.contains(n) => {
                    let class = __w.module_class_name.clone();
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
        // `ref.null $t` → a WASM GC typed null (traps on struct.get/array.get
        // per spec), distinct from a plain null. Lowered to the compiler builtin
        // `__wast_typed_null` which emits `ref.null` (Op::NULL) with a non-zero
        // heap-type immediate so the VM produces a `TypedNull`.
        // ⛔ THE HEAPTYPE IS THE INSTRUCTION. This passed `vec![]` — the
        // immediate was parsed and then THROWN AWAY, so every `ref.null`
        // spelling collapsed to one intrinsic that emits `HT_NONE`. In the
        // VM that is invisible (one null), but `none` bottoms the ANY
        // hierarchy and `noextern` bottoms the EXTERN one, so `nullref` is
        // NOT a subtype of `externref`: `(ref.null extern)` stored into an
        // externref global came out `ref.null none` and V8 answered
        // `global.set[0] expected type externref, found ref.null of type
        // nullref`. Found by Fathom's 10-line descriptor module.
        //
        // ⚠ `parse_vt` (the stack-typing path, `"ref.null"` below) reads this
        // same immediate CORRECTLY. So the checker believed `externref` while
        // the emitter wrote `none` — the two halves of this walker disagreed,
        // which is exactly why no fixture could catch it.
        //
        // Concrete `$T` keeps the old lowering: `ref.null $T` is in the ANY
        // hierarchy either way, and resolving the index needs the compiler's
        // type table, which this arm does not have.
        "ref.null" => {
            // ⚠ THE IMMEDIATE ARRIVES AS A STRING, NOT AN IDENT.
            // `instr_arg_inner_to_expr` maps `Rule::bare_heap_type` to
            // `Expression::string(..)`; only `$id` becomes an `Ident`. Reading
            // just one shape is how this looked fixed and stayed broken.
            let spelling = args.first().and_then(|e| match &e.kind {
                ExprKind::Lit(Literal::Str(t)) => abs_heap(t.trim()),
                ExprKind::Ident(n) => abs_heap(n),
                _ => None,
            });
            let arg = match spelling {
                Some(ht) => vec![Expression::string(ht)],
                None => vec![],
            };
            Ok(make_call("__wast_typed_null", arg, span))
        }
        // ref.func $f → a first-class reference to module function `$f`. Module
        // functions are static methods of the module class, so this is the
        // static method referenced as a value (the compiler tears it off into a
        // funcref). ref.func by numeric index is not resolved here (needs the
        // compiler's chunk table); named refs cover the common case.
        "ref.func" => {
            let field = match args.into_iter().next() {
                Some(e) => match &e.kind {
                    ExprKind::Ident(n) => n.clone(),
                    _ => return Ok(e),
                },
                None => return Ok(Expression::null()),
            };
            // ⚠ AN IMPORTED FUNCTION IS NOT A METHOD OF THIS MODULE'S CLASS.
            //
            // This always qualified with the CURRENT module's class, so
            // `ref.func $1` where `$1` is `(import "A" "f" (func $1 …))`
            // produced `<ThisModule>.$1` — a member that does not exist —
            // and the reference came out `undefined`. `call $1` worked the
            // whole time because the CALL path resolves through
            // `import_alias`; only the reference-taking path did not, so the
            // two spellings of "use this imported function" disagreed.
            //
            // `import_alias` already holds the EXPORTER's (class, method),
            // resolved through `(register …)` or a component instantiation.
            let (class, field) = match __w.import_alias.get(&field) {
                Some((exporter_class, method)) => (exporter_class.clone(), method.clone()),
                None => (__w.module_class_name.clone(), field),
            };
            Ok(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::ident(&class)),
                    field,
                    null_safe: false,
                },
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
        // ── Call Tags proposal ────────────────────────────────────────────
        // `call_with_tag $tag : [ti* funcref] -> [to*]` — same operand layout
        // as `call_ref` (funcref on top), plus the tag immediate that says
        // which CONVENTION the call is made under.
        "call_with_tag" | "call_return_with_tag" => {
            let mut rest = args;
            let tag = if rest.is_empty() {
                String::new()
            } else {
                wasm_type_ref_name(&rest.remove(0))
            };
            let callee = rest.pop().unwrap_or_else(Expression::null);
            let tag = __w.call_tag_alias.get(&tag).cloned().unwrap_or(tag);
            Ok(Expression::with_span(
                ExprKind::WasmCallWithTag {
                    tag,
                    callee: Box::new(callee),
                    args: rest,
                    tail: name == "call_return_with_tag",
                    table: None,
                },
                span,
            ))
        }
        // "`call_indirect_with_tag $table $call_tag` is shorthand for
        // `(call_with_tag $call_tag (table.get $table))`" — so it desugars to
        // exactly that, and there is one path rather than two to keep in step.
        "call_indirect_with_tag" => {
            let mut rest = args;
            let table = if rest.is_empty() {
                Expression::float(0.0)
            } else {
                rest.remove(0)
            };
            let tag = if rest.is_empty() {
                String::new()
            } else {
                wasm_type_ref_name(&rest.remove(0))
            };
            let elem_index = rest.pop().unwrap_or_else(Expression::null);
            let table_idx = match &table.kind {
                ExprKind::Lit(Literal::Int(i)) => *i as u32,
                ExprKind::Ident(n) => resolve_table_index(__w, n) as u32,
                _ => 0,
            };
            Ok(Expression::with_span(
                ExprKind::WasmCallWithTag {
                    tag,
                    // The element index, not a funcref — `table` says so.
                    callee: Box::new(elem_index),
                    args: rest,
                    tail: false,
                    table: Some(table_idx),
                },
                span,
            ))
        }

        "call_ref" | "return_call_ref" => build_call_ref_opcode(__w, &name, args, span),

        // ── GC / WasmGC struct ops ────────────────────────────────────────
        // struct.new $T v0 v1 ...  → {"0": v0, "1": v1, ..., "__type": "T"}
        // args: [typeidx, field_val_0, field_val_1, ...]. The `__type` stamp
        // carries the GC type name so the VM's `ref.test`/`ref.cast`/`br_on_cast`
        // resolve identity + subtyping through the registered type hierarchy.
        "struct.new" => {
            let type_name = resolve_wast_type_name(__w, args.first());
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
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_type(obj, &type_name, span))
        }
        // struct.new_default $T → each field set to its storage type's default
        // (0 for ints, 0.0 for floats, null for refs), stamped with its rtt.
        "struct.new_default" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let field_types = __w.struct_field_types.get(&type_name).cloned()
                .unwrap_or_default();
            let props: Vec<ObjectProperty> = field_types
                .iter()
                .enumerate()
                .map(|(i, ty)| ObjectProperty::KeyValue {
                    key: Expression::string(&i.to_string()),
                    value: default_value_for_storage_type(ty),
                })
                .collect();
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_type(obj, &type_name, span))
        }
        // ── Custom Descriptors allocation ─────────────────────────────────
        // struct.new_desc $T v0 .. vN-1 desc → the same positional object
        // `struct.new` builds, plus the descriptor. args:
        // [typeidx, field_val_0 .. field_val_N-1, descriptor] — the descriptor
        // is LAST (Overview.md §"Allocation With Descriptors").
        "struct.new_desc" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let rest = args.get(1..).unwrap_or(&[]);
            let (vals, desc): (Vec<Expression>, Expression) = match rest.split_last() {
                Some((d, v)) => (v.to_vec(), d.clone()),
                None => (vec![], Expression::with_span(ExprKind::Lit(Literal::Null), span)),
            };
            let props: Vec<ObjectProperty> = vals
                .into_iter()
                .enumerate()
                .map(|(i, v)| ObjectProperty::KeyValue {
                    key: Expression::string(&i.to_string()),
                    value: v,
                })
                .collect();
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_desc_type(obj, &type_name, desc, span))
        }
        // struct.new_default_desc $T desc → every field at its default, plus
        // the descriptor. args: [typeidx, descriptor].
        "struct.new_default_desc" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let desc = args
                .get(1)
                .cloned()
                .unwrap_or_else(|| Expression::with_span(ExprKind::Lit(Literal::Null), span));
            let field_types = __w
                .struct_field_types
                .get(&type_name)
                .cloned()
                .unwrap_or_default();
            let props: Vec<ObjectProperty> = field_types
                .iter()
                .enumerate()
                .map(|(i, ty)| ObjectProperty::KeyValue {
                    key: Expression::string(&i.to_string()),
                    value: default_value_for_storage_type(ty),
                })
                .collect();
            let obj = Expression::with_span(ExprKind::Object(props), span);
            Ok(wast_stamp_desc_type(obj, &type_name, desc, span))
        }
        // ref.get_desc $T → the descriptor attached at allocation.
        // args: [typeidx, ref].
        "ref.get_desc" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let target = args
                .get(1)
                .cloned()
                .unwrap_or_else(|| Expression::with_span(ExprKind::Lit(Literal::Null), span));
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_ref_get_desc")),
                    args: vec![
                        Argument::positional(target),
                        Argument::positional(Expression::string(&type_name)),
                    ],
                    optional: false,
                },
                span,
            ))
        }
        // array.new_default $T → a length-N array filled with the element type's
        // default. For numeric elements that's 0/0.0, so lower to `array.new $T
        // <default> <length>` (which fills the value); ref elements keep the VM's
        // null-fill via `array_new_default`. args: [typeidx, length].
        "array.new_default" => {
            let elem = args.first().and_then(|a| match &a.kind {
                ExprKind::Ident(n) => __w.array_elem_type.get(&qualify_type_name(__w, n)).cloned(),
                _ => None,
            });
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
            let type_name = resolve_wast_type_name(__w, args.first());
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
            let arr = Expression::with_span(ExprKind::Array(vals), span);
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("__wast_stamp_array_type")),
                    args: vec![
                        Argument::positional(arr),
                        Argument::positional(Expression::string(&type_name)),
                    ],
                    optional: false,
                },
                span,
            ))
        }
        // struct.get $T N ref  → ref["N"] (null-trapped). The `_s`/`_u` variants
        // sign/zero-extend a packed i8/i16 field, mirroring array.get_s/get_u.
        // args: [typeidx, fieldidx, ref_expr]
        "struct.get" | "struct.get_s" | "struct.get_u" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let field_idx = resolve_struct_field_index(__w, &type_name, args.get(1));
            let obj = args.into_iter().nth(2).unwrap_or(Expression::null());
            let member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(obj),
                    field: field_idx.to_string(),
                    null_safe: false,
                },
                span,
            );
            if name == "struct.get" {
                return Ok(member);
            }
            let signed = name == "struct.get_s";
            let field_ty = {
                __w.struct_field_types
                    .get(&type_name)
                    .and_then(|v| v.get(field_idx as usize).cloned())
            };
            Ok(match (field_ty.as_deref(), signed) {
                (Some("i8"), true) => make_call("i32.extend8_s", vec![member], span),
                (Some("i16"), true) => make_call("i32.extend16_s", vec![member], span),
                (Some("i8"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(member),
                    right: Box::new(Expression::int(0xFF)),
                }),
                (Some("i16"), false) => Expression::new(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(member),
                    right: Box::new(Expression::int(0xFFFF)),
                }),
                // Non-packed (i32/i64/ref/…): the stored value is already exact.
                _ => member,
            })
        }
        // struct.set $T N ref val → ref["N"] = val  (produces null, used as stmt)
        // args: [typeidx, fieldidx, ref_expr, val_expr]
        "struct.set" => {
            let type_name = resolve_wast_type_name(__w, args.first());
            let field_idx = resolve_struct_field_index(__w, &type_name, args.get(1));
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

fn walk_instr_arg_pair(__w: &mut WastWalker, pair: Pair<Rule>, labels: &mut LabelStack) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty instr_arg")?;
    match inner.as_rule() {
        Rule::folded_instr => walk_folded_instr_as_expr(__w, inner, Span::default(), labels),
        // `(ref null? ht)` used to reach here as a `folded_instr` with `ref`
        // as its head. It now has its own rule, so rebuild the same shape the
        // heap-type resolver reads: a call to `ref` whose args carry the
        // optional `null` marker and the heap type (an `$id` as an ident, a
        // spec spelling as a string).
        Rule::ref_type_arg => {
            let mut args: Vec<Argument> = Vec::new();
            let text = inner.as_str();
            if text
                .split_whitespace()
                .any(|t| t.trim_start_matches('(') == "null")
            {
                args.push(Argument::positional(Expression::ident("null")));
            }
            // ⚠ `(exact $x)` IS A HEAP TYPE, and it does not parse as one.
            //
            // The grammar spells `(ref null? (exact_heap_type | heap_type))`,
            // so with `(ref null (exact $super))` the child is
            // `exact_heap_type` and a `find(heap_type)` matches NOTHING. The
            // name was then simply absent from the lowered call, and every
            // cast against an exact type compared against the EMPTY string —
            // `trap: ref.cast_null failed: value is not ` with nothing after
            // "is not" is that empty name surfacing.
            //
            // `exact_heap_type` wraps an `index`, so unwrap one level and the
            // rest of the resolver is unchanged.
            let exact = inner
                .clone()
                .into_inner()
                .any(|c| c.as_rule() == Rule::exact_heap_type);
            if exact {
                // A marker argument, the same shape as `null` above — the
                // resolver reads markers by name, so exactness reaches the
                // compiler without a new argument position to keep in sync.
                args.push(Argument::positional(Expression::ident("exact")));
            }
            let ht = inner
                .clone()
                .into_inner()
                .find(|c| c.as_rule() == Rule::heap_type)
                .or_else(|| {
                    inner
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::exact_heap_type)
                        .and_then(|e| e.into_inner().next())
                });
            if let Some(ht) = ht {
                let s = ht.as_str();
                args.push(Argument::positional(match s.strip_prefix('$') {
                    // A `$id` names a MODULE-LOCAL type, so it carries the same
                    // module qualifier the declaration site applies — otherwise
                    // `ref.test (ref $t)` resolves against whichever module
                    // declared `$t` last. Spec spellings (`any`, `struct`, …)
                    // are abstract and stay bare.
                    Some(name) => Expression::ident(&qualify_type_name(__w, name)),
                    None => Expression::string(s),
                }));
            }
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("ref")),
                args,
                optional: false,
            }))
        }
        _ => Ok(instr_arg_inner_to_expr(inner)),
    }
}

/// True when an instruction's FLOAT immediate is stored at SINGLE precision.
///
/// It matters at PARSE time and nowhere afterwards. `f32.const` lowers to
/// `f32_demote_f64` over an f64 literal, so a significand longer than 24 bits
/// is rounded TWICE — once on the way into the f64, again on the demote — and
/// the spec's rounding section is made of exactly the literals where those two
/// answers differ. `+0x1.00000100000000001p-50` sits just ABOVE the f32 tie and
/// must round up; rounding to f64 first lands it exactly ON the tie, where
/// ties-to-even sends it back down to `0x1.000000p-50`. Reading it straight to
/// f32 sees the bit that decides it; widening the result back to f64 is exact,
/// so the demote that follows is a no-op.
fn instr_float_is_f32(name: &str) -> bool {
    name == "f32.const"
}

/// Convert one raw `instr_arg`, reading a float immediate at the width its
/// instruction actually stores it at.
fn walk_instr_arg_for(
    __w: &mut WastWalker,
    raw: Pair<Rule>,
    labels: &mut LabelStack,
    name: &str,
) -> Result<Expression, String> {
    if instr_float_is_f32(name)
        && let Some(inner) = raw.clone().into_inner().next()
        && inner.as_rule() == Rule::float
    {
        return Ok(parse_float_at(inner.as_str(), true));
    }
    walk_instr_arg_pair(__w, raw, labels)
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

/// The binding name a global lowers to: its `$id`, or a synthetic
/// `__wasm_global_<module>_<idx>` when unnamed, so `global.get <idx>` can
/// resolve to it and two modules' unnamed globals stay distinct — they share
/// one top-level binding space, and `__wasm_global_0` from a second module used
/// to shadow the first's, making `(get $m1 "g")` read $m2's value.
fn global_binding_name(__w: &mut WastWalker, pair: &Pair<Rule>, idx: usize) -> String {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
        .unwrap_or_else(|| {
            let m = __w.current_module_seq;
            format!("__wasm_global_{m}_{idx}")
        })
}

fn walk_global_field(__w: &mut WastWalker, pair: Pair<Rule>, idx: usize) -> Result<(String, Expression), String> {
    let name = global_binding_name(__w, &pair, idx);
    let mut init = Expression::int(0);
    let mut labels = LabelStack::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::instr {
            init = walk_instr_as_expr(__w, child, &mut labels)?;
        }
    }
    Ok((name, init))
}

// ── Linear memory + data segments ─────────────────────────────────────────────

/// Walk a `(memory …)` field. Returns the declaration plus, for the inline
/// `(memory (data "…"))` abbreviation, the active data segment it stands for:
/// the memory is sized to exactly hold the data (⌈len/64Ki⌉ pages, min and max
/// both), and the bytes land at offset 0.
fn walk_memory_field(__w: &mut WastWalker, pair: Pair<Rule>, memory_idx: usize) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    // An IMPORTED memory that resolved to an exporter's memory declares
    // NOTHING: it is that memory, and its own written limits are a link-time
    // constraint rather than a size. Declaring one anyway is what gave the two
    // modules separate bytes — and what made `memory_size_import` answer the
    // import's minimum (1) instead of the exporter's actual size (2).
    if let Some((_, true)) = __w.memory_field_info.get(memory_idx).copied() {
        return Ok(Vec::new());
    }
    let mut min_pages: u64 = 0;
    let mut max_pages: Option<u64> = None;
    let mut inline_bytes: Option<Vec<u8>> = None;
    // `(memory i64 …)` — memory64's 64-bit index space. It appears inside
    // `mem_type`, and, in the `(memory i64 (data …))` abbreviation, as a direct
    // child of the field.
    let mut is_64 = false;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::addr_type => is_64 = child.as_str() == "i64",
            Rule::mem_type => {
                is_64 = child
                    .clone()
                    .into_inner()
                    .any(|p| p.as_rule() == Rule::addr_type && p.as_str() == "i64");
                let mut nums = child.into_inner().filter(|p| p.as_rule() == Rule::integer);
                if let Some(min) = nums.next() {
                    min_pages = parse_wat_u64(min.as_str());
                }
                if let Some(max) = nums.next() {
                    max_pages = Some(parse_wat_u64(max.as_str()));
                }
            }
            Rule::inline_memory_data => {
                let mut bytes: Vec<u8> = Vec::new();
                for s in child.into_inner().filter(|p| p.as_rule() == Rule::string) {
                    bytes.extend(decode_wat_data_string(s.as_str()));
                }
                inline_bytes = Some(bytes);
            }
            _ => {}
        }
    }

    let Some(bytes) = inline_bytes else {
        return Ok(vec![Statement::with_span(
            StmtKind::MemoryDecl {
                min_pages,
                max_pages,
                is_64,
            },
            span,
        )]);
    };

    const PAGE: u64 = 65536;
    let pages = (bytes.len() as u64).div_ceil(PAGE);
    // `(memory (data …))` declares THIS field's own memory, so the segment
    // targets whatever slot this module's memidx `memory_idx` resolved to.
    let own_slot = __w
        .memory_field_info
        .get(memory_idx)
        .map(|(s, _)| *s)
        .unwrap_or(__w.memory_index_base + memory_idx) as u32;
    Ok(vec![
        Statement::with_span(
            StmtKind::MemoryDecl {
                min_pages: pages,
                max_pages: Some(pages),
                is_64,
            },
            span,
        ),
        Statement::with_span(
            StmtKind::DataSegment {
                memory_index: own_slot,
                offset: Some(Expression::int(0)),
                bytes,
            },
            span,
        ),
    ])
}

/// The method name a function `index` (`$id` or a positional integer) refers
/// to, given the module's index→name map. `None` when a positional index is out
/// of range (a malformed module, caught by validation, not by silent fallback).
/// The funcref an element-segment item denotes.
///
/// ⚠ TWO resolutions, both of which were missing. An item written as a NUMBER
/// (`(elem (table $t) (i32.const 2) func 3 1 4 1)`) names a function INDEX, not
/// a method called "3" — left verbatim it built a member access on a name no
/// class has, and the slot silently stayed null. And an index in the imported
/// range denotes another module's function, which is not a member of THIS
/// module's class at all; it resolves through `import_alias` exactly as a
/// direct call to the same import does.
///
/// `bulk-memory/table_init.wast` needs both at once: five imported functions
/// followed by an element list written entirely in numbers.
fn elem_item_funcref(__w: &WastWalker, item: &str, class: &str) -> Expression {
    match elem_item_owner(__w, item, class) {
        Some((owner, method)) => Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&owner)),
            field: method,
            null_safe: false,
        }),
        None => Expression::null(),
    }
}

/// The `(owner class, method)` an element-segment item resolves to, shared by
/// the active path (which builds a member access) and the passive one (which
/// hands the pair to the compile-time registration).
fn elem_item_owner(__w: &WastWalker, item: &str, class: &str) -> Option<(String, String)> {
    if item.is_empty() {
        return None;
    }
    let name = match item.parse::<usize>() {
        Ok(i) => __w.func_index_name.get(i).cloned()?,
        Err(_) => item.to_string(),
    };
    Some(match __w.import_alias.get(&name) {
        Some((cls, m)) => (cls.clone(), m.clone()),
        None => (class.to_string(), name),
    })
}

/// The binding an UNNAMED standalone `(import "m" "e" (func …))` is reached by.
///
/// The index pre-scan ALREADY gives such an import a slot under this name — it
/// has occupied the leading function indices all along. What was missing is the
/// other half: `import_alias` is keyed by LOCAL ID, an unnamed import has none,
/// so nothing ever bound the name to the exporter's method and a numeric
/// reference resolved to an undefined identifier. Imports precede definitions
/// (WASM 3.0 §6.4), so an import's ordinal among function imports IS its
/// function index, and the two pre-scans agree without either seeing the other.
fn imported_func_binding_name(func_index: usize) -> String {
    format!("__wasm_import_{func_index}")
}

fn resolve_func_index_name(idx: &Pair<Rule>, index_names: &[String]) -> Option<String> {
    let inner = idx.clone().into_inner().next()?;
    match inner.as_rule() {
        Rule::id => Some(inner.as_str()[1..].to_string()),
        Rule::integer => index_names
            .get(parse_wat_u64(inner.as_str()) as usize)
            .cloned(),
        _ => None,
    }
}

/// Extract `(tag $e (param t*))`'s name (without `$`) and payload arity.
/// The two module/name strings of an `(import "m" "n" …)` field.
fn scan_import_names(pair: &Pair<Rule>) -> Option<(String, String)> {
    let strings: Vec<String> = pair
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::string)
        .map(|c| unquote(c.as_str()))
        .collect();
    match strings.as_slice() {
        [m, n, ..] => Some((m.clone(), n.clone())),
        _ => None,
    }
}

/// A tag declaration's id, payload arity, inline export names, and inline
/// import target. The identity of a tag is its DECLARATION, so an import (which
/// declares nothing — it names someone else's tag) has to be told apart from a
/// definition here, not at the reference sites.
#[allow(clippy::type_complexity)]
/// `(call_tag $t (param …) (result …) (fallback $f)?)` — Call Tags proposal.
///
/// Records the tag's arity in `call_tag_params` as it goes: the folder needs it
/// to know how many operands `call_with_tag` consumes, and a tag can be used
/// before its declaration is walked.
fn walk_call_tag_field(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name: Option<String> = None;
    let mut params = 0usize;
    let mut results = 0usize;
    let mut signature = String::new();
    let mut fallback: Option<String> = None;
    let mut canonical = false;
    let mut imported: Option<String> = None;
    let mut exported: Vec<String> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            // Strip the `$`: an instruction's `$t` arrives as `Ident("t")`, so
            // keeping the sigil here would make the declaration and its uses
            // two different names — and the call site would mint a second,
            // empty tag rather than finding this one.
            Rule::id => name = Some(child.as_str().trim_start_matches('$').to_string()),
            Rule::typeuse => {
                let (p, r) = count_typeuse_params_results(__w, &child);
                params = p;
                results = r;
                signature = typeuse_signature(__w, &child);
            }
            Rule::call_tag_canon => canonical = true,
            Rule::call_tag_fallback => {
                if let Some(idx) = child.into_inner().find(|c| c.as_rule() == Rule::index) {
                    fallback = Some(idx.as_str().trim_start_matches('$').to_string());
                }
            }
            // "Similarly, one can import and export call tags." An IMPORTED tag
            // keeps the exporter's identity, which is what makes the proposal's
            // security property meaningful: a module that exports none of its
            // tags cannot have its funcs reached indirectly from outside.
            // Resolution is by NAME, exactly as exception-tag imports resolve,
            // so the two modules meet at one entity.
            Rule::import_inline => {
                let names: Vec<String> = child
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::string)
                    .map(|c| unquote(c.as_str()))
                    .collect();
                if names.len() == 2 {
                    imported = Some(format!("{}:{}", names[0], names[1]));
                }
            }
            Rule::export_inline => {
                if let Some(sp) = child.into_inner().find(|c| c.as_rule() == Rule::string) {
                    exported.push(unquote(sp.as_str()));
                }
            }
            _ => {}
        }
    }
    let name = name.unwrap_or_else(|| format!("#call_tag{}", __w.call_tag_params.len()));
    __w.call_tag_params.insert(name.clone(), (params, results));
    // Both directions key the SAME entity name so an import and the export it
    // resolves to are one tag — identity is the whole contract.
    if let Some(ref ext) = imported {
        __w.call_tag_params.insert(ext.clone(), (params, results));
    }
    for e in &exported {
        __w.call_tag_params.insert(e.clone(), (params, results));
        // Register the tag under its export name as well, so an importing
        // module naming that export resolves to THIS entity.
        __w.call_tag_alias.insert(name.clone(), e.clone());
    }
    Ok(Statement::with_span(
        StmtKind::WasmCallTagDecl {
            // An imported tag is the EXPORTER's entity: key by the external
            // name so both modules resolve to one id.
            // The declared `$id` stays the key either way — the call site
            // names the tag by it. `canonical` only changes which ENTITY that
            // name resolves to: the signature-interned one, or a fresh one.
            // Entity name, in priority order: an IMPORT takes the exporter's
            // identity; an EXPORT publishes under the export name so an
            // importer naming it finds this entity; otherwise the local `$id`.
            name: imported
                .or_else(|| exported.first().cloned())
                .unwrap_or(name),
            params: params as u8,
            results: results as u8,
            signature,
            canonical,
            fallback,
        },
        span,
    ))
}

/// `(import "m" "e" (call_tag $t (param …) (result …)))` — the outer spelling.
///
/// Keyed by `m:e`, the EXTERNAL name, exactly as the inline
/// `(call_tag $t (import "m" "e") …)` form is, so the two spellings and the
/// exporting module's declaration all resolve to one entity.
fn walk_imported_call_tag(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let strings: Vec<String> = pair
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::string)
        .map(|c| unquote(c.as_str()))
        .collect();
    // The EXPORT NAME is the shared identity: a wast script publishes exports
    // by name, so the importing module's tag and the exporting module's are the
    // same entity precisely when they agree on it.
    let external = strings.get(1).cloned().unwrap_or_default();
    let mut local = String::new();
    let mut params = 0usize;
    let mut results = 0usize;
    let mut signature = String::new();
    if let Some(desc) = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::import_desc)
    {
        for child in desc.into_inner() {
            match child.as_rule() {
                Rule::id => local = child.as_str().trim_start_matches('$').to_string(),
                Rule::typeuse => {
                    let (p, r) = count_typeuse_params_results(__w, &child);
                    params = p;
                    results = r;
                    signature = typeuse_signature(__w, &child);
                }
                _ => {}
            }
        }
    }
    // The LOCAL alias is what this module's `call_with_tag` names, so it must
    // resolve to the imported entity — register both spellings.
    if !local.is_empty() {
        __w.call_tag_params.insert(local.clone(), (params, results));
        __w.call_tag_alias.insert(local, external.clone());
    }
    __w.call_tag_params.insert(external.clone(), (params, results));
    Ok(Statement::with_span(
        StmtKind::WasmCallTagDecl {
            name: external,
            params: params as u8,
            results: results as u8,
            signature,
            canonical: false,
            fallback: None,
        },
        span,
    ))
}

/// `(func_switch $s (case $tag $func)* (forward $other)?)`.
fn walk_func_switch_field(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name: Option<String> = None;
    let mut arms: Vec<(String, String)> = Vec::new();
    let mut forward: Option<String> = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str().trim_start_matches('$').to_string()),
            Rule::func_switch_arm => {
                let idxs: Vec<String> = child
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::index)
                    .map(|c| c.as_str().trim_start_matches('$').to_string())
                    .collect();
                if idxs.len() == 2 {
                    arms.push((idxs[0].clone(), idxs[1].clone()));
                }
            }
            Rule::func_switch_forward => {
                if let Some(idx) = child.into_inner().find(|c| c.as_rule() == Rule::index) {
                    forward = Some(idx.as_str().trim_start_matches('$').to_string());
                }
            }
            _ => {}
        }
    }
    let _ = &__w;
    Ok(Statement::with_span(
        StmtKind::WasmFuncSwitchDecl {
            name: name.unwrap_or_default(),
            arms,
            forward,
        },
        span,
    ))
}

/// Parameter and result counts of a `typeuse`, for a call tag's signature.
/// The typeuse's DECLARED VALUE TYPES, as a signature string (`"i32,i32->f64"`).
///
/// ⛔ THE TYPES WERE BEING COUNTED AND THROWN AWAY.
/// `count_typeuse_params_results` below reduces `(param i32) (result i32)` to
/// `(1, 1)`, and everything downstream — `StmtKind::WasmCallTagDecl`,
/// `CallTagDef`, `canonical_call_tags: HashMap<(u8,u8), u32>` — keys on that
/// pair. So `call_tag.canon [i32]->[i32]` and `call_tag.canon [f64]->[f64]`
/// intern to the SAME tag, and an `i32`-shaped funcref is callable under the
/// `f64` canonical tag. Measured: that module is accepted today.
///
/// The Overview says `call_tag.canon $functype` derives the canonical tag *of
/// that functype*; two functypes are two tags. And since `call_indirect $table
/// $functype` is now shorthand for `call_with_tag (call_tag.canon $functype)`,
/// the Security property — "a funcref called under a convention it does not
/// handle STOPS, rather than being called anyway under the wrong shape" — is
/// only ARITY-safety while the key is a pair of counts.
///
/// A string rather than a type model on purpose: the VM erases runtime types
/// (`Chunk` carries `arity`/`param_count`/`result_arity` and no value types at
/// all), so the declared spelling is the only place the functype survives. It
/// is compared, never interpreted.
/// The CANONICAL name of an instruction's `(type $t)`, when it names one.
///
/// ⛔ THE IDENTITY, NOT THE SHAPE. `call_indirect`'s runtime check must tell
/// `$f1` from `$f2` when both are `(func)` in DIFFERENT rec groups — the
/// signature strings are equal there and only the canonicalised name differs.
/// `qualify_type_name` runs the same canonicalisation the callee's
/// `declared_func_type` went through, so the two are comparable.
///
/// Empty when the instruction spells its type inline (`(param …)(result …)`),
/// which has no name to compare — the signature check covers that case.
fn typeuse_canon_name(__w: &WastWalker, pair: &Pair<Rule>) -> String {
    let host = match pair.as_rule() {
        Rule::instr => pair.clone().into_inner().next().unwrap_or_else(|| pair.clone()),
        _ => pair.clone(),
    };
    for c in host.into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        let Some(bt) = c.into_inner().next() else { continue };
        if bt.as_rule() != Rule::block_type {
            continue;
        }
        if !bt.as_str().trim_start_matches('(').trim_start().starts_with("type") {
            continue;
        }
        if let Some(ix) = bt.into_inner().find(|x| x.as_rule() == Rule::index) {
            let raw = ix.as_str().trim_start_matches('$').to_string();
            return match raw.parse::<usize>() {
                Ok(i) => __w.type_index_name.get(i).cloned().unwrap_or(raw),
                Err(_) => qualify_type_name(__w, &raw),
            };
        }
    }
    String::new()
}

fn typeuse_signature(__w: &mut WastWalker, pair: &Pair<Rule>) -> String {
    let mut params: Vec<String> = Vec::new();
    let mut results: Vec<String> = Vec::new();
    // Set when a `(type $t)` names a functype this walker cannot resolve; the
    // caller records nothing rather than a fabricated shape.
    let mut unresolved = false;

    // ⛔ ON AN INSTRUCTION THE TYPE IS A `block_type`, NOT A `typeuse`. This
    // read `Rule::typeuse`/`param`/`result`/`index` off the pair's children and
    // found NONE of them, so every `call_indirect (type $t)` produced `"->"` —
    // a real signature meaning "no params, no results", not an absent one. The
    // runtime then had nothing true to compare and fell back to arity, which
    // cannot tell `(func (result i32))` from `(func (result i64))`.
    //
    // `peek_typeuse_shape` reads the same instruction correctly and is the
    // shape mirrored here: `instr_arg` → `block_type`, whose KEYWORD is inlined
    // by the grammar and so has to be read off the node's own text.
    let host = match pair.as_rule() {
        Rule::instr => pair.clone().into_inner().next().unwrap_or_else(|| pair.clone()),
        _ => pair.clone(),
    };
    let mut saw_block_type = false;
    for c in host.clone().into_inner() {
        if c.as_rule() != Rule::instr_arg {
            continue;
        }
        let Some(bt) = c.into_inner().next() else { continue };
        if bt.as_rule() != Rule::block_type {
            continue;
        }
        saw_block_type = true;
        let head = bt.as_str().trim_start_matches('(').trim_start();
        let vals = |b: &Pair<Rule>| -> Vec<String> {
            b.clone()
                .into_inner()
                .filter(|x| matches!(x.as_rule(), Rule::val_type | Rule::any_val_type))
                .map(|x| {
                    canonical_val_type(
                        __w,
                        &x.as_str().split_whitespace().collect::<Vec<_>>().join(" "),
                    )
                })
                .collect()
        };
        if head.starts_with("type") {
            let named = bt
                .clone()
                .into_inner()
                .find(|x| x.as_rule() == Rule::index)
                .map(|x| x.as_str().trim_start_matches('$').to_string());
            match named
                .map(|n| qualify_type_name(__w, &n))
                .and_then(|n| __w.type_func_sigs.get(&n).cloned())
            {
                Some((p, r)) => {
                    params.extend(p);
                    results.extend(r);
                }
                None => unresolved = true,
            }
        } else if head.starts_with("param") {
            params.extend(vals(&bt));
        } else if head.starts_with("result") {
            results.extend(vals(&bt));
        }
    }

    // A real `typeuse` node (the shape a func FIELD carries) still reads the
    // old way; only the instruction spelling goes through `block_type`.
    if !saw_block_type {
        for child in host.into_inner() {
            match child.as_rule() {
                Rule::param => {
                    for t in child.into_inner().filter(|c| {
                        matches!(c.as_rule(), Rule::any_val_type | Rule::val_type)
                    }) {
                        params.push(canonical_val_type(
                            __w,
                            &t.as_str().split_whitespace().collect::<Vec<_>>().join(" "),
                        ));
                    }
                }
                Rule::result => {
                    for t in child.into_inner().filter(|c| {
                        matches!(c.as_rule(), Rule::any_val_type | Rule::val_type)
                    }) {
                        results.push(canonical_val_type(
                            __w,
                            &t.as_str().split_whitespace().collect::<Vec<_>>().join(" "),
                        ));
                    }
                }
                Rule::index => {
                    let n = qualify_type_name(__w, child.as_str().trim_start_matches('$'));
                    match __w.type_func_sigs.get(&n) {
                        Some((p, r)) => {
                            params.extend(p.iter().cloned());
                            results.extend(r.iter().cloned());
                        }
                        None => unresolved = true,
                    }
                }
                _ => {}
            }
        }
    }

    if unresolved {
        return String::new();
    }
    format!("{}->{}", params.join(","), results.join(","))
}

fn count_typeuse_params_results(__w: &mut WastWalker, pair: &Pair<Rule>) -> (usize, usize) {
    let mut params = 0usize;
    let mut results = 0usize;
    for child in pair.clone().into_inner() {
        match child.as_rule() {
            Rule::param => {
                params += child
                    .into_inner()
                    .filter(|c| matches!(c.as_rule(), Rule::any_val_type | Rule::val_type))
                    .count()
                    .max(1);
            }
            Rule::result => {
                results += child
                    .into_inner()
                    .filter(|c| matches!(c.as_rule(), Rule::any_val_type | Rule::val_type))
                    .count()
                    .max(1);
            }
            // `(type $t)` — take the referenced type's shape.
            Rule::index => {
                // ⚠ TWO bugs here: the `$` was kept, so this never matched even
                // before type names were module-qualified; and the key must now
                // carry the module qualifier like every other type lookup.
                let n = qualify_type_name(__w, child.as_str().trim_start_matches('$'));
                if let Some(p) = __w.type_func_params.get(&n).copied() {
                    params = p;
                }
            }
            _ => {}
        }
    }
    (params, results)
}

fn scan_tag_decl(__w: &mut WastWalker,
    pair: &Pair<Rule>,
) -> (Option<String>, u8, Vec<String>, Option<(String, String)>) {
    let (name, arity) = scan_tag_signature(__w, pair.clone());
    let mut exports: Vec<String> = Vec::new();
    let mut import: Option<(String, String)> = None;
    for child in pair.clone().into_inner() {
        match child.as_rule() {
            Rule::export_inline => {
                if let Some(s) = child.into_inner().find(|c| c.as_rule() == Rule::string) {
                    exports.push(unquote(s.as_str()));
                }
            }
            Rule::import_inline => import = scan_import_names(&child),
            _ => {}
        }
    }
    (name, arity, exports, import)
}

fn scan_tag_signature(__w: &mut WastWalker, pair: Pair<Rule>) -> (Option<String>, u8) {
    let mut name: Option<String> = None;
    let mut arity: u8 = 0;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::id => name = Some(child.as_str()[1..].to_string()),
            Rule::tag_type => {
                // A tagtype is a typeuse: `(type $t)`, an inline `(param …)*`,
                // or both. Inline params win when present (they are the spec's
                // explicit restatement of the referenced type); otherwise the
                // arity comes from the named func type.
                let mut inline = 0usize;
                let mut type_ref: Option<String> = None;
                for p in child.into_inner() {
                    match p.as_rule() {
                        Rule::param => {
                            inline += p
                                .into_inner()
                                .filter(|v| v.as_rule() == Rule::any_val_type)
                                .count();
                        }
                        Rule::index => {
                            type_ref = Some(p.as_str().trim_start_matches('$').to_string());
                        }
                        _ => {}
                    }
                }
                arity = if inline > 0 {
                    inline as u8
                } else {
                    type_ref
                        .map(|n| qualify_type_name(__w, &n))
                        .and_then(|n| __w.type_func_params.get(&n).copied())
                        .unwrap_or(0) as u8
                };
            }
            _ => {}
        }
    }
    (name, arity)
}

/// `(tag $e (param t*))` — an exception-tag declaration. Emits a `WasmTagDecl`
/// the compiler imports as a load-time tag entity. Arities are recorded in the
/// module pre-scan (so `catch $e` sees them regardless of source order).
fn walk_tag_field(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let (name, arity) = scan_tag_signature(__w, pair);
    let ordinal = {
        let cur = __w.tag_decl_ordinal;
        __w.tag_decl_ordinal += 1;
        cur
    };
    Ok(Statement::with_span(
        StmtKind::WasmTagDecl {
            // Same entity name the pre-scan registered, so a numeric reference,
            // an anonymous declaration and an import alias all meet at one
            // entity.
            name: tag_ref_name(__w, &name.unwrap_or_else(|| format!("#{ordinal}"))),
            arity,
        },
        span,
    ))
}

/// Walk a `(table …)` field. Returns the table declaration (goes BEFORE the
/// module class, since the VM allocates the table at instantiation) and, for the
/// inline `(table t (elem $f …))` abbreviation, its active-segment population
/// (goes AFTER the class — it references the funcs as static methods, so it must
/// run once the class exists, exactly like a standalone `(elem …)` field).
fn walk_table_field(__w: &mut WastWalker, 
    pair: Pair<Rule>,
    table_idx: usize,
) -> Result<(Statement, Vec<Statement>), String> {
    let span = to_span(&pair);
    let mut min_size: u64 = 0;
    let mut max_size: Option<u64> = None;
    let mut has_table_type = false;
    // `(table i64 …)` — a 64-bit table index space (memory64).
    let mut is_64 = false;
    // Inline `(table t (elem $f …))` abbreviation: the `index*` funcidx list.
    let mut inline_funcs: Vec<String> = Vec::new();
    // WASM 3.0 table INIT EXPRESSION — `(table $t 10 funcref (ref.func $d))`
    // fills EVERY slot with the value, instead of the type's default null.
    let mut init_expr: Option<Expression> = None;
    let mut labels = LabelStack::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::folded_instr => {
                init_expr = Some(walk_folded_instr_as_expr(__w, child, span, &mut labels)?);
            }
            // `(table $t i64 funcref (elem …))` — the abbreviation states the
            // index type outside `table_type`.
            Rule::addr_type => is_64 = child.as_str() == "i64",
            Rule::table_type => {
                // table_type = addr_type? integer integer? ref_type — the index
                // type, then min and optional max.
                has_table_type = true;
                is_64 = child
                    .clone()
                    .into_inner()
                    .any(|p| p.as_rule() == Rule::addr_type && p.as_str() == "i64");
                let mut nums = child.into_inner().filter(|p| p.as_rule() == Rule::integer);
                if let Some(min) = nums.next() {
                    min_size = parse_wat_u64(min.as_str());
                }
                if let Some(max) = nums.next() {
                    max_size = Some(parse_wat_u64(max.as_str()));
                }
            }
            // ⛔ A NUMERIC funcidx IS NOT A NAME. `(table funcref (elem 0 1))`
            // is the same index space `call 0` uses — pushing the literal "0"
            // as a function NAME matched nothing, so the slot stayed null and
            // the first `call_indirect` trapped "uninitialized element 0".
            // The named spelling worked, which is why this survived: only
            // func_ptrs.wast writes the numeric one.
            Rule::index => {
                let raw = child.as_str().trim();
                let resolved = match raw.strip_prefix('$') {
                    Some(n) => n.to_string(),
                    None => raw
                        .parse::<usize>()
                        .ok()
                        .and_then(|i| __w.func_index_name.get(i).cloned())
                        .unwrap_or_else(|| raw.to_string()),
                };
                inline_funcs.push(resolved);
            }
            _ => {}
        }
    }

    // Inline elem abbreviation ≡ a table sized to the element count plus an
    // active elem segment populating it from slot 0.
    if !has_table_type && !inline_funcs.is_empty() {
        let n = inline_funcs.len() as u64;
        let class = __w.module_class_name.clone();
        let decl = Statement::with_span(
            StmtKind::TableDecl {
                min_size: n,
                max_size: Some(n),
                is_64,
            },
            span,
        );
        let mut population = Vec::new();
        for (i, f) in inline_funcs.iter().enumerate() {
            // ⛔ AN IMPORTED FUNCTION IS NOT A MEMBER OF THIS MODULE'S CLASS.
            // `(table funcref (elem $print_i32))` names an IMPORT, and
            // qualifying it as `ThisModule.print_i32` resolves to nothing — the
            // slot stayed null and the first `call_indirect` trapped
            // "uninitialized element 0". The `call` arm has always resolved
            // these through `import_alias` (an import is a second name for the
            // EXPORTING module's method) and `host_import_alias`; the inline
            // elem list never did.
            let funcref = if let Some((icls, imeth)) = __w.import_alias.get(f).cloned() {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&icls)),
                    field: imeth,
                    null_safe: false,
                })
            } else if let Some(hostname) = __w.host_import_alias.get(f).cloned() {
                // A HOST import is reached by its bare `host:m:n` ident, not
                // through any class.
                Expression::ident(&hostname)
            } else {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class)),
                    field: f.clone(),
                    null_safe: false,
                })
            };
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

    // A table INIT EXPRESSION fills the whole table at instantiation. Emitted as
    // ONE `table.fill`, never `min_size` separate `table.set`s — a table may be
    // declared with billions of slots and the statement list would be the
    // program. Operand order is the VM's: (dst, value, count) after the tableidx
    // immediate. `min_size == 0` needs no fill at all.
    let population = match init_expr {
        Some(v) if min_size > 0 => vec![Statement::with_span(
            StmtKind::Expr(make_call(
                "table_fill",
                vec![
                    Expression::int(table_idx as i64),
                    Expression::int(0),
                    v,
                    Expression::int(min_size as i64),
                ],
                span,
            )),
            span,
        )],
        _ => Vec::new(),
    };
    Ok((
        Statement::with_span(StmtKind::TableDecl { min_size, max_size, is_64 }, span),
        population,
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
/// The composite type of a `(type …)` field, as whitespace-normalised TEXT.
///
/// This is the structural identity WASM 3.0 canonicalisation compares, and it
/// is read from the source rather than rebuilt from the parsed shape because
/// the parsed shape is lossy in exactly the place that matters: field and
/// element MUTABILITY. `(struct (field (mut i32)))` and `(struct (field i32))`
/// are different types, and both parse to a storage type of "i32".
///
/// Whitespace is collapsed so two spellings of the same type agree; nothing
/// else is normalised, so a difference we cannot interpret leaves the types
/// DISTINCT — which is the status quo, not a wrong merge.
fn composite_type_text(type_field: &Pair<Rule>) -> String {
    fn find(p: &Pair<Rule>) -> Option<String> {
        for c in p.clone().into_inner() {
            if matches!(
                c.as_rule(),
                Rule::composite_type
                    | Rule::struct_type
                    | Rule::array_type
                    | Rule::func_type
            ) {
                return Some(c.as_str().to_string());
            }
            if let Some(t) = find(&c) {
                return Some(t);
            }
        }
        None
    }
    find(type_field)
        .unwrap_or_default()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

/// The declared NAME a type immediate refers to: an `$id` verbatim, a positional
/// index mapped through declaration order.
///
/// `struct.get_s 0 0` names type 0. Left as the literal `"0"`, every per-type
/// lookup missed — including the packed field's storage type, so `get_s` did no
/// sign extension and answered 254 where the spec says -2.
/// A wast type name, qualified by the module that declares it.
///
/// Each module has its own type index space; the compiler's type table is one
/// script-wide, name-keyed store. Without a qualifier the two disagree and the
/// LAST declaration of a name silently redefines every earlier one.
///
/// Abstract heap types (`any`, `struct`, `func`, `eq`, …) and the already
/// qualified/synthetic names are left alone — only a module-local `$id` gets a
/// prefix. Empty stays empty so "no type" keeps meaning no type.
fn qualify_type_name(__w: &WastWalker, name: &str) -> String {
    if name.is_empty() || name.starts_with("m#") {
        return name.to_string();
    }
    if vybe_runtime::opcode::heaptype::HeapType::from_spec_name(name).is_some() {
        return name.to_string();
    }
    let qualified = format!("m#{}#{}", __w.current_module_seq, name);
    // One funnel, so `struct.new`, `ref.cast`, field lookups and diagnostics
    // all agree on which row a `$id` names. Empty during the pre-scan that
    // BUILDS the map, which is what keeps that pass reading raw names.
    __w.type_canonical
        .get(&qualified)
        .cloned()
        .unwrap_or(qualified)
}

fn resolve_wast_type_name(__w: &WastWalker, expr: Option<&Expression>) -> String {
    let raw = expr.map(wasm_type_ref_name).unwrap_or_default();
    match raw.parse::<usize>() {
        // `type_index_name` already holds QUALIFIED names — it is built from
        // the declaration site — so a numeric reference needs no further work.
        Ok(i) => __w
            .type_index_name
            .get(i)
            .cloned()
            .unwrap_or_else(|| qualify_type_name(__w, &raw)),
        Err(_) => qualify_type_name(__w, &raw),
    }
}

fn wasm_type_ref_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(n) => n.clone(),
        ExprKind::Lit(Literal::Int(i)) => i.to_string(),
        _ => String::new(),
    }
}

/// Stamp a freshly-built struct object with its WASM GC rtt (the registered
/// type's id): `__wast_stamp_type(obj, "T")` → the compiler emits
/// `GLOBAL_GET __tid_T` + `SET_TYPE_ID`, so the instance carries the real
/// `type_id` the VM's `ref.test`/`ref.cast`/`is_subtype` read — no `__type`
/// string. This is the struct analogue of `array.new`'s rtt stamp.
/// `struct.new_desc` / `struct.new_default_desc` — as [`wast_stamp_type`], with
/// the descriptor carried as a third argument so the compiler can push it on
/// top of the field values.
fn wast_stamp_desc_type(
    obj: Expression,
    type_name: &str,
    descriptor: Expression,
    span: Span,
) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident("__wast_stamp_desc_type")),
            args: vec![
                Argument::positional(obj),
                Argument::positional(Expression::string(type_name)),
                Argument::positional(descriptor),
            ],
            optional: false,
        },
        span,
    )
}

fn wast_stamp_type(obj: Expression, type_name: &str, span: Span) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident("__wast_stamp_type")),
            args: vec![
                Argument::positional(obj),
                Argument::positional(Expression::string(type_name)),
            ],
            optional: false,
        },
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
/// An explicit `(table $t)` target resolves through `resolve_table_index`.
/// `None` in `funcs` is a NULL element expression: it occupies its slot without
/// storing anything.
fn walk_elem_field(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut offset: i64 = 0;
    // Default target is THIS module's table 0 — program table <base> in a
    // multi-module script.
    let mut table_index: i64 = __w.table_index_base as i64;
    let mut funcs: Vec<Option<String>> = Vec::new();
    // Index-aligned with `funcs`: the walked element EXPRESSION for any item
    // that is not a plain `ref.func`/`ref.null`. Both lists advance together so
    // a mixed segment keeps every slot in its declared position.
    let mut item_exprs: Vec<Option<Expression>> = Vec::new();
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
                        table_index = resolve_table_index(__w, &tname);
                    }
                    is_active = true;
                }
            }
            Rule::index => funcs.push(Some(child.as_str().trim_start_matches('$').to_string())),
            // `(item (ref.func $f))` or a bare `(ref.func $f)` element: the
            // initializer is a `ref.func` whose funcidx is the first id/index
            // inside. A `(ref.null …)` element has no funcref — but it STILL
            // OCCUPIES ITS SLOT. Dropping it shifted every later entry down one:
            // elem.wast:164 `(elem (i32.const 6) funcref (ref.null func)
            // (ref.func $a))` put `$a` at slot 6 instead of 7. A funcref const
            // expression is only `ref.func`/`ref.null`/`global.get`, so the
            // presence of the mnemonic identifies the null form.
            Rule::elem_item => {
                let text = child.as_str();
                if text.contains("ref.null") {
                    funcs.push(None);
                    item_exprs.push(None);
                } else if text.contains("ref.func") {
                    funcs.push(first_ident_or_index(&child));
                    item_exprs.push(None);
                } else {
                    // WASM 3.0 element segments hold arbitrary element
                    // EXPRESSIONS of their reftype, not only funcrefs:
                    // `(item (ref.i31 (i32.const 999)))`, `(item (global.get
                    // $g))`, `(item (struct.new $t …))`. Reducing every item to
                    // a function NAME dropped all of them — the slot stayed at
                    // the table's default and the read answered null.
                    let mut labels = LabelStack::new();
                    let mut expr: Option<Expression> = None;
                    for sub in child.clone().into_inner() {
                        match sub.as_rule() {
                            Rule::instr => expr = Some(walk_instr_as_expr(__w, sub, &mut labels)?),
                            Rule::folded_instr => {
                                expr =
                                    Some(walk_folded_instr_as_expr(__w, sub, span, &mut labels)?)
                            }
                            _ => {}
                        }
                    }
                    funcs.push(None);
                    item_exprs.push(expr);
                }
            }
            _ => {}
        }
    }
    // Every element segment (active / passive / declarative) occupies one slot
    // in the element index space, in declaration order.
    let seg_index = {
        let i = __w.elem_seg_counter;
        __w.elem_seg_counter = i + 1;
        __w.elem_index_base + i
    };
    if !is_active {
        if is_declare {
            // Declarative: only permits `ref.func`, no runtime payload.
            return Ok(Statement::with_span(StmtKind::Block(Vec::new()), span));
        }
        // Passive: register the element list under this segment index so a later
        // `table.init $e` / `array.new_elem $e` copies real funcrefs from it.
        // Compile-time directive resolved to function chunk indices; the VM
        // materializes the funcrefs at instantiation (see `passive_elem_funcs`).
        //
        // ⚠ The CLASS is an argument. Each module in a script gets its own
        // (`__wasm_module`, `__wasm_module_1`, …), and the consumer used to
        // resolve names under a hardcoded `__wasm_module` — so every passive
        // segment after the FIRST module resolved to nothing, was stored empty,
        // and `table.init` reported "missing element segment".
        let class = __w.module_class_name.clone();
        let mut args = vec![Expression::int(seg_index as i64)];
        // ⚠ ONE ARGUMENT PER SLOT, AND IT IS AN EXPRESSION.
        //
        // An element segment holds ELEMENT EXPRESSIONS (WASM 3.0 §4.5.4), so
        // each slot sends the expression the VM evaluates at instantiation. A
        // `ref.func` item sends `__elem_func(owner, method)` instead, because
        // resolving a function name to a chunk is the compiler's job — and the
        // OWNER matters, since an item may name an IMPORTED function, which
        // lives in the exporting module's class rather than this one's.
        //
        // Emitting only the resolvable items renumbered the segment: an
        // `(elem funcref (ref.null func) (ref.func $a))` put `$a` at index 0,
        // so `table.init` copying one element from offset 1 read past the end.
        for (i, f) in funcs.iter().enumerate() {
            match f {
                Some(item) => {
                    let (owner, method) = elem_item_owner(__w, item, &class)
                        .unwrap_or_else(|| (String::new(), String::new()));
                    args.push(make_call(
                        "__elem_func",
                        vec![Expression::string(&owner), Expression::string(&method)],
                        span,
                    ));
                }
                // `ref.null` carries no expression and lands as a null slot; a
                // general element expression sends itself.
                None => args.push(
                    item_exprs
                        .get(i)
                        .cloned()
                        .flatten()
                        .unwrap_or_else(Expression::null),
                ),
            }
        }
        return Ok(Statement::with_span(
            StmtKind::Expr(make_call("__wast_register_passive_elem", args, span)),
            span,
        ));
    }
    let class = __w.module_class_name.clone();
    let mut stmts = Vec::new();
    for (i, f) in funcs.iter().enumerate() {
        // A null element leaves its slot at the table's default — the slot is
        // still consumed, which is why the index comes from `enumerate` and not
        // from a counter that only advances on real funcrefs.
        let funcref = match f {
            Some(f) => elem_item_funcref(__w, f, &class),
            // A general element expression fills the slot with its value.
            None => match item_exprs.get(i).cloned().flatten() {
                Some(e) => e,
                None => continue,
            },
        };
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
            optional: false,
        }),
        _ => Expression::null(),
    }
}

/// The ordered field storage types of a `struct_type`/`struct_subtype` body.
/// A defined function's signature as VALUE TYPES. Inline `(param …)`/
/// `(result …)` win when present — they are the spec's explicit restatement
/// of the type — otherwise it comes from the `(type $t)` the function names.
/// Mirrors how `scan_tag_signature` resolves a tag's arity.
/// The function type a `(func $f (type $t) …)` field DECLARES, module-qualified.
///
/// Distinct from its structural signature: two types with identical params and
/// results are the same STRUCTURE but different NAMES, and Custom Descriptors'
/// exact casts discriminate on the name.
fn func_field_declared_type(__w: &WastWalker, func_field: &Pair<Rule>) -> Option<String> {
    for c in func_field.clone().into_inner() {
        if c.as_rule() != Rule::typeuse {
            continue;
        }
        for t in c.into_inner() {
            if t.as_rule() == Rule::index {
                let raw = t.as_str().trim_start_matches('$').to_string();
                return Some(match raw.parse::<usize>() {
                    Ok(i) => __w.type_index_name.get(i).cloned().unwrap_or(raw),
                    Err(_) => qualify_type_name(__w, &raw),
                });
            }
        }
    }
    None
}

/// A value type with every `$type` reference in it replaced by its CANONICAL
/// name.
///
/// ⛔ A SIGNATURE IS COMPARED, SO ITS SPELLINGS MUST AGREE. `call_indirect`
/// checks the callee's signature against the call site's, and both are built
/// from source text: `(func (param (ref $s1)))` and `(func (param (ref $s2)))`
/// are the SAME type when `$s1` and `$s2` are, but their texts differ, so the
/// comparison rejected a valid call. Running the reference through the same
/// funnel every other consumer uses makes the two sides say the same thing.
///
/// Only `$`-prefixed references are touched — `i32`, `funcref`, `(ref null
/// any)` and the rest have no name to resolve and pass through unchanged.
fn canonical_val_type(__w: &WastWalker, text: &str) -> String {
    if !text.contains('$') {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find('$') {
        out.push_str(&rest[..at]);
        let tail = &rest[at + 1..];
        let end = tail
            .find(|c: char| !(c.is_alphanumeric() || "_.+-*/\\^~=<>!?@#$%&|:'`".contains(c)))
            .unwrap_or(tail.len());
        if end == 0 {
            out.push('$');
            rest = tail;
            continue;
        }
        out.push_str(&qualify_type_name(__w, &tail[..end]));
        rest = &tail[end..];
    }
    out.push_str(rest);
    out
}

fn func_field_signature(__w: &WastWalker, func_field: &Pair<Rule>) -> (Vec<String>, Vec<String>) {
    let mut params: Vec<String> = Vec::new();
    let mut results: Vec<String> = Vec::new();
    let mut type_ref: Option<String> = None;
    for c in func_field.clone().into_inner() {
        if c.as_rule() != Rule::typeuse {
            continue;
        }
        for t in c.into_inner() {
            match t.as_rule() {
                Rule::index => {
                    type_ref = Some(t.as_str().trim_start_matches('$').to_string())
                }
                Rule::param | Rule::result => {
                    let target = if t.as_rule() == Rule::param {
                        &mut params
                    } else {
                        &mut results
                    };
                    for v in t.into_inner() {
                        if matches!(v.as_rule(), Rule::any_val_type | Rule::val_type) {
                            target.push(canonical_val_type(
                                __w,
                                &v.as_str().split_whitespace().collect::<Vec<_>>().join(" "),
                            ));
                        }
                    }
                }
                _ => {}
            }
        }
    }
    if params.is_empty() && results.is_empty() {
        // The maps are keyed by the QUALIFIED name (the declaration site
        // qualifies), while a `(type $t)` typeuse yields the bare `$t`.
        if let Some(sig) = type_ref
            .map(|n| qualify_type_name(__w, &n))
            .and_then(|n| __w.type_func_sigs.get(&n))
        {
            return sig.clone();
        }
    }
    (params, results)
}

/// The value types of a `(func (param …)* (result …)*)` composite, as the
/// spec's own spellings.
///
/// A function type's identity is STRUCTURAL — `Comptype_sub/func` matches
/// `(FUNC t_11* -> t_12*)` against `(FUNC t_21* -> t_22*)` by matching the
/// parameter and result types, with no name anywhere in the rule. Comparing
/// arities alone would make `(func (param i32))` and `(func (param f64))`
/// indistinguishable, so the types themselves are what gets recorded.
///
/// Returned as `(params, results)`, each in declaration order. A single
/// `(param i32 i64)` declares two parameters, so the val types are flattened
/// rather than counted per clause.
fn func_type_signature(composite_inner: &Pair<Rule>) -> (Vec<String>, Vec<String>) {
    let mut params: Vec<String> = Vec::new();
    let mut results: Vec<String> = Vec::new();
    for p in composite_inner.clone().into_inner() {
        let target = match p.as_rule() {
            Rule::param => &mut params,
            Rule::result => &mut results,
            _ => continue,
        };
        for v in p.into_inner() {
            if matches!(v.as_rule(), Rule::any_val_type | Rule::val_type) {
                target.push(v.as_str().split_whitespace().collect::<Vec<_>>().join(" "));
            }
        }
    }
    (params, results)
}

fn struct_field_types(composite_inner: &Pair<Rule>) -> Vec<String> {
    composite_inner
        .clone()
        .into_inner()
        .filter(|p| p.as_rule() == Rule::field_def)
        .flat_map(|f| field_storage_types(&f))
        .collect()
}

/// Every storage type ONE `field_def` declares, in order.
///
/// `(field i8 i16 i32)` is the abbreviation for THREE consecutive unnamed
/// fields, and `(field)` for none — a NAMED field declares exactly one. Reading
/// only the first shifted every later field's index, so `struct.get $T 2` on
/// such a struct addressed the wrong slot.
fn field_storage_types(field_def: &Pair<Rule>) -> Vec<String> {
    field_def
        .clone()
        .into_inner()
        .filter(|c| matches!(c.as_rule(), Rule::storage_type | Rule::ref_val_type))
        .map(|c| array_elem_type(&c).unwrap_or_else(|| c.as_str().to_string()))
        .collect()
}

/// The field NAMES, index-aligned with `struct_field_types`: a named field's
/// `$id` without the sigil, and the empty string for every field of an
/// unnamed abbreviation.
///
/// `struct.get $T $y` addresses a field BY NAME. Without this the name reached
/// the lowering as an `Ident`, failed the "is it an integer literal" test, and
/// fell back to index **0** — self-consistently, so a `set $y` / `get $y` pair
/// round-tripped and only a MIXED numeric/named pair (`gc/struct.wast`'s
/// `set_get_1`) ever showed it.
/// The field index a `struct.get`/`struct.set` immediate names: a literal
/// index verbatim, a `$id` looked up in the struct's declared field names.
///
/// A name that resolves to nothing stays 0 — the same fallback as before, so a
/// struct the prescan never saw behaves as it used to rather than shifting.
fn resolve_struct_field_index(
    __w: &WastWalker,
    type_name: &str,
    field_arg: Option<&Expression>,
) -> i64 {
    match field_arg.map(|a| &a.kind) {
        Some(ExprKind::Lit(Literal::Int(i))) => *i,
        Some(ExprKind::Ident(name)) => __w
            .struct_field_ids
            .get(type_name)
            .and_then(|ids| ids.iter().position(|f| f == name))
            .map(|i| i as i64)
            .unwrap_or(0),
        _ => 0,
    }
}

fn struct_field_names(composite_inner: &Pair<Rule>) -> Vec<String> {
    let mut out = Vec::new();
    for field in composite_inner
        .clone()
        .into_inner()
        .filter(|p| p.as_rule() == Rule::field_def)
    {
        let id = field
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::id)
            .map(|c| c.as_str().trim_start_matches('$').to_string());
        let count = field_storage_types(&field).len();
        match id {
            Some(name) if count == 1 => out.push(name),
            _ => out.extend(std::iter::repeat_n(String::new(), count)),
        }
    }
    out
}

/// Resolve a table reference (`$t1` name or a numeric index) to its table index.
fn resolve_table_index(__w: &mut WastWalker, name: &str) -> i64 {
    // Named tables were registered ALREADY shifted into the script's index
    // space; a numeric index is module-relative and shifts here. A name that is
    // neither is the module's default table — also base-relative.
    let base = __w.table_index_base as i64;
    __w.table_name_index.get(name).copied()
        .map(|i| i as i64)
        .or_else(|| name.parse::<i64>().ok().map(|n| base + n))
        .unwrap_or(base)
}

fn walk_data_field(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    // Default and numeric memory indices are MODULE-RELATIVE; `memory_slots`
    // maps them onto the script's index space, and an IMPORTED memory's entry
    // points at the exporter's slot so a segment written here lands in the
    // memory the other module can read.
    let mut memory_index: u32 = default_memory_slot(__w) as u32;
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
                                    let n = parse_wat_u64(i.as_str()) as usize;
                                    memory_index = __w
                                        .memory_slots
                                        .get(n)
                                        .copied()
                                        .unwrap_or(__w.memory_index_base + n)
                                        as u32;
                                } else {
                                    // `(memory $m2)` — resolve the name to its
                                    // declaration index.
                                    let name = i.as_str().trim_start_matches('$');
                                    memory_index = __w.memory_name_index.get(name).copied()
                                        .unwrap_or(0)
                                        as u32;
                                }
                            }
                        }
                        Rule::instr => offset = Some(walk_instr_as_expr(__w, m, &mut labels)?),
                        Rule::folded_instr => {
                            let sp = to_span(&m);
                            offset = Some(walk_folded_instr_as_expr(__w, m, sp, &mut labels)?);
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

fn walk_invoke_cmd(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut func_name = String::new();
    let mut module_id: Option<String> = None;
    let mut args: Vec<Expression> = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            // `(invoke $M "e" …)` targets the module NAMED $M; a bare
            // `(invoke "e" …)` targets the most recent one.
            Rule::id => module_id = Some(child.as_str()[1..].to_string()),
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
    let resolved = match &module_id {
        Some(m) => __w.module_exports.get(m).and_then(|e| e.get(&func_name).cloned())
            .map(|method| (m.clone(), method)),
        None => __w.export_func_map.get(&func_name).cloned()
            .map(|method| (__w.module_class_name.clone(), method)),
    };
    let callee = match resolved {
        Some((class, method)) => Expression::with_span(
            ExprKind::Member {
                object: Box::new(Expression::ident(&class)),
                field: method,
                null_safe: false,
            },
            span,
        ),
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
/// The FAILURE condition for a reference expectation: `ref_is_null(v)` must be
/// 1 when null was expected and 0 when a non-null reference was. Reference
/// results are never compared with `!=` — a typed null is a GC reference, and
/// scalar equality against it is not a null-ness test.
fn ref_null_check(value: Expression, want_null: bool, span: Span) -> Expression {
    Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(make_call("ref_is_null", vec![value], span)),
            right: Box::new(Expression::int(if want_null { 0 } else { 1 })),
        },
        span,
    )
}

/// The FAILURE condition for ONE expected result: true exactly when `actual`
/// does not match what this `result_val` asks for.
///
/// Every shape answers the same question, which is what lets `(either …)` be
/// expressed here at all: an implementation conforms if ANY alternative
/// matches, so the `either` failure is that EVERY alternative failed.
fn result_failure_cond(
    child: &Pair<Rule>,
    actual: Expression,
    span: Span,
) -> Result<Expression, String> {
    let text = child.as_str().trim();

    // `(either r1 r2 …)` — the relaxed-simd proposal states results the
    // implementation may CHOOSE between (fused vs unfused multiply-add, the
    // NaN it propagates). Answering any of them is conforming.
    if text.starts_with("(either") {
        let mut fail: Option<Expression> = None;
        for alt in child
            .clone()
            .into_inner()
            .filter(|c| c.as_rule() == Rule::result_val)
        {
            let cond = result_failure_cond(&alt, actual.clone(), span)?;
            fail = Some(match fail {
                None => cond,
                Some(prev) => Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(prev),
                        right: Box::new(cond),
                    },
                    span,
                ),
            });
        }
        // `(either)` with nothing in it constrains nothing, so it cannot fail.
        return Ok(fail.unwrap_or_else(|| Expression::int(0)));
    }

    // The payload-less reference shapes: `(ref.array)` says "the result is a
    // non-null reference to an array", not "the result equals something".
    // That is precisely `ref.test` against the abstract heap type — the
    // NON-nullable form, so a null reference fails it.
    let bare = text.trim_start_matches('(').trim_end_matches(')').trim();
    if let Some(heap) = match bare {
        "ref.array" => Some("array"),
        "ref.struct" => Some("struct"),
        "ref.i31" => Some("i31"),
        "ref.eq" => Some("eq"),
        "ref.any" => Some("any"),
        "ref.exn" => Some("exn"),
        _ => None,
    } {
        return Ok(Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(make_call(
                    "ref_test",
                    vec![Expression::string(heap), actual],
                    span,
                )),
                right: Box::new(Expression::int(0)),
            },
            span,
        ));
    }

    // `(ref.null …)` = a null reference; bare `(ref.func)`, `(ref.extern)` and
    // `(ref.host)` = SOME non-null reference. A payload-carrying
    // `(ref.extern N)` / `(ref.host N)` falls through to the value compare.
    if text.contains("ref.null") {
        return Ok(ref_null_check(actual, true, span));
    }
    if matches!(bare, "ref.func" | "ref.extern" | "ref.host") {
        return Ok(ref_null_check(actual, false, span));
    }

    let want = walk_const_expr(child.clone())?;

    // `nan:canonical` / `nan:arithmetic` pin no payload, so they cannot be
    // compared with `==`. NaN is the only value that differs from itself, so
    // "equals itself" is precisely the failure case.
    if text.contains("nan") {
        return Ok(Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(actual.clone()),
                right: Box::new(actual),
            },
            span,
        ));
    }

    // A v128 is a VECTOR: scalar `!=` does not compare its lanes, so it
    // reports "different" for two identical vectors and every v128-returning
    // assert fails regardless of its lanes. Lane-wise instead: `i8x16.eq`
    // yields all-ones in each byte lane that matches and `i8x16.all_true` is 1
    // only when every lane did, so the FAILURE is that result being 0.
    // Comparing at i8x16 is shape-independent — it checks all 16 bytes, so it
    // is correct for i32x4 / f64x2 / any other reading of the same bits.
    if text.contains("v128") {
        return Ok(Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(make_call(
                    "i8x16.all_true",
                    vec![make_call("i8x16.eq", vec![actual, want], span)],
                    span,
                )),
                right: Box::new(Expression::int(0)),
            },
            span,
        ));
    }

    Ok(Expression::with_span(
        ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(actual),
            right: Box::new(want),
        },
        span,
    ))
}

fn walk_assert_return(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    // The expected results, kept as PARSE PAIRS: every shape's comparison is
    // built by `result_failure_cond`, which needs the source form (a value
    // compare, a null-ness predicate, a `ref.test`, a lane-wise vector
    // compare, or an `either` over any of those) and not just a value.
    let mut expected: Vec<Pair<Rule>> = Vec::new();
    // ⛔ "assert_return failed" ALONE NAMES NOTHING. A file stops at its first
    // failing assertion, so a bare message says only that SOMETHING in a
    // 200-assertion file disagreed — not which, and not what it wanted. Both
    // halves are STATIC TEXT already in the fixture, so quoting them costs no
    // runtime plumbing and turns "this file fails" into a diagnosis.
    let mut action_text = String::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action => {
                action_text = squeeze_ws(child.as_str());
                action_expr = Some(walk_action(__w, child)?);
            }
            Rule::result_val => expected.push(child),
            _ => {}
        }
    }
    let want_text = |ps: &[Pair<Rule>]| -> String {
        ps.iter()
            .map(|p| squeeze_ws(p.as_str()))
            .collect::<Vec<_>>()
            .join(" ")
    };
    let Some(action) = action_expr else {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    };

    if expected.len() == 1 {
        let msg = format!(
            "assert_return failed: {} — expected {}",
            action_text,
            want_text(&expected)
        );
        let want = expected.pop().expect("checked len");
        let cond = result_failure_cond(&want, action, span)?;
        let throw = Statement::with_span(
            StmtKind::Throw {
                expr: Some(Expression::string(&msg)),
                cause: None,
            },
            span,
        );
        return Ok(Statement::with_span(
            StmtKind::If {
                cond,
                then_body: vec![throw],
                elifs: Vec::new(),
                else_body: None,
            },
            span,
        ));
    }

    if expected.is_empty() {
        // `(assert_return (invoke …))` with no results: the assertion is
        // exactly that the action completes without trapping — run it.
        return Ok(Statement::with_span(StmtKind::Expr(action), span));
    }

    // Multi-value: the callee's multi-value ABI packs the results as one
    // tuple/array value (see `multi_value_return_stmt`), so bind it to a
    // temp and compare element-wise — the same NaN ("equals itself" is the
    // failure) and v128 (i8x16.eq + all_true) semantics as the
    // single-value path, index by index.
    let tmp = "__wast_mv";
    let decl = Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp.to_string()),
                type_hint: None,
                init: Some(action),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        },
        span,
    );
    let elem = |i: usize| {
        Expression::with_span(
            ExprKind::Index {
                object: Box::new(Expression::ident(tmp)),
                index: Box::new(Expression::int(i as i64)),
                null_safe: false,
            },
            span,
        )
    };
    let mut fail: Option<Expression> = None;
    for (i, want) in expected.iter().enumerate() {
        let cond = result_failure_cond(want, elem(i), span)?;
        fail = Some(match fail {
            None => cond,
            Some(prev) => Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(prev),
                    right: Box::new(cond),
                },
                span,
            ),
        });
    }
    let throw = Statement::with_span(
        StmtKind::Throw {
            expr: Some(Expression::string(&format!(
                "assert_return failed: {} — expected {}",
                action_text,
                want_text(&expected)
            ))),
            cause: None,
        },
        span,
    );
    let check = Statement::with_span(
        StmtKind::If {
            cond: fail.expect("expected values are non-empty here"),
            then_body: vec![throw],
            elifs: Vec::new(),
            else_body: None,
        },
        span,
    );
    Ok(Statement::with_span(
        StmtKind::Block(vec![decl, check]),
        span,
    ))
}

/// `(assert_exception action)` — the action must raise an exception; the
/// directive names no tag and compares no message, so ANY escape is the pass.
///
/// This cannot reuse `assert_trap`'s language-level `try`: that lowers to a
/// `try_table` with one TYPED clause for the language's own `vybe:exception`
/// tag, and a spec exception carries a `(tag …)` entity that never matches it —
/// the throw escaped the assertion as an uncaught runtime error. The clause
/// that matches "any tag" is `catch_all`, so the assertion is built out of the
/// spec node directly.
///
/// Normal completion sets a flag INSIDE the try body; `catch_all` branches past
/// it, leaving the flag false. The failure throw therefore sits after the whole
/// `try_table`, where the handler can no longer swallow it.
/// `(assert_malformed (module quote "…") "message")` — the quoted text must
/// FAIL TO PARSE.
///
/// This is the one member of the invalid/malformed/unlinkable trio that needs
/// no validator: malformedness is a property of the TEXT, and the grammar is
/// right here. All three used to lower to `Empty`, so the whole validation half
/// of the spec suite passed vacuously.
///
/// The `(module binary "…")` form is a DECODE assertion about a byte string,
/// not a parse assertion about text — a different service, so it stays skipped
/// rather than being answered wrongly.
///
/// A leniency in our grammar surfaces here as a red test rather than as
/// silence, which is the point: text the spec calls malformed and we accept is
/// a real finding, and this is what reports it.
/// Text-level malformities that survive the grammar, checked over a parse tree.
///
/// The grammar matches an instruction name with a generic dotted-identifier
/// rule — deliberately, so a new mnemonic needs no grammar edit — and it
/// matches an integer literal without knowing which width will consume it.
/// Both are the right trade for ordinary compilation and wrong for
/// `assert_malformed`, which is precisely a claim about the text. So the two
/// are asked HERE, of the quoted module only:
///
///   * an instruction name that resolves to no opcode (`invalid.opcode`),
///   * an `iN.const` whose literal does not fit N bits ("constant out of
///     range"), and
///   * an import/export NAME whose bytes are not valid UTF-8. A name is a
///     character string and must decode; a DATA string is a byte string and
///     `\ff` there is an ordinary byte, so the check is on names only.
/// Structural VALIDITY checks — the properties a module that PARSES can still
/// violate. Returns the spec's diagnostic for the first violation found.
///
/// ⛔ VALIDITY IS NOT MALFORMEDNESS. `align=7` is malformed text (not a power
/// of two); `align=8` on a one-byte load parses perfectly and is INVALID. The
/// suite asserts them with different commands and different messages, so a
/// check in the wrong half makes the other half's assertion pass for the wrong
/// reason.
///
/// This is the structural subset only. The official suite carries 2720
/// `assert_invalid` assertions and **2297 of them are "type mismatch"** — the
/// stack-typing algorithm with block signatures, which does not exist yet
/// ([[the validator]]). What is implemented here are the checks that need no
/// type system at all:
///
///   * `alignment must not be larger than natural` — 99 assertions
///   * `invalid lane index`                        — 48
///   * `duplicate export name`                     — 17
///
/// Everything else still returns None, which means the assertion FAILS rather
/// than passing quietly. That is the point: an unimplemented check must not
/// look like a satisfied one.
fn module_invalid_reason(pairs: pest::iterators::Pairs<Rule>) -> Option<String> {
    use std::collections::HashSet;
    let mut export_names: HashSet<String> = HashSet::new();
    let pairs: Vec<Pair<Rule>> = pairs.collect();
    for pair in pairs.clone() {
        if let Some(r) = module_invalid_walk(pair, &mut export_names) {
            return Some(r);
        }
    }
    // ⛔ STACK TYPING RUNS LAST, AND THAT ORDER IS THE SPEC'S. Presence before
    // agreement: a module with `local.get 99` AND a type error must report
    // "unknown local". Every structural rule above settles first, and the
    // one-directional message comparison turns any ordering slip into a
    // visible wrong-message failure rather than a silent pass.
    for pair in &pairs {
        if let Some(r) = stack_typing_reason_in(pair) {
            return Some(r);
        }
    }
    None
}

/// Name/index resolution over one module: every reference must name something
/// the module declares. These need no type system — they are the second-biggest
/// group in the suite after "type mismatch".
///
/// ⛔ ONLY UNAMBIGUOUS OPERAND POSITIONS ARE CHECKED. `memory.init $d` and
/// `i32.load $m` carry an operand whose entity kind depends on how many bare
/// indices are written (the memidx peel), and guessing wrong reports the wrong
/// "unknown X" — which the message comparison now turns into a visible failure
/// rather than a silent pass. Where the position is not certain, nothing is
/// reported and the assertion stays honestly red.
/// `global.set` on an immutable global. ⛔ The machinery for this already
/// existed in `validate_module`, whose errors are routed to the MALFORMED
/// check — so the suite's `assert_invalid "immutable global"` never saw it.
/// Writing to a `const` global is a typing property of a module that parses
/// perfectly: invalid, not malformed.
fn immutable_global_reason(module: &Pair<Rule>) -> Option<String> {
    let mut immut: std::collections::HashSet<String> = std::collections::HashSet::new();
    for pair in module.clone().into_inner() {
        let inner = match pair.as_rule() {
            Rule::module_field => match pair.clone().into_inner().next() {
                Some(i) => i,
                None => continue,
            },
            _ => pair.clone(),
        };
        if inner.as_rule() != Rule::global_field {
            continue;
        }
        let children: Vec<_> = inner.clone().into_inner().collect();
        let is_mut = children
            .iter()
            .any(|c| c.as_rule() == Rule::global_type && c.as_str().contains("mut"));
        if !is_mut {
            if let Some(id) = census_id(&inner) {
                immut.insert(id);
            }
        }
    }
    if immut.is_empty() {
        return None;
    }
    let mut targets = Vec::new();
    collect_global_set_targets(module.clone(), &mut targets);
    targets
        .into_iter()
        .find(|t| immut.contains(t))
        .map(|t| format!("immutable global: ${t}"))
}

/// §4.4.7: a memarg's `offset` must fit the memory's ADDRESS WIDTH. On a
/// 32-bit memory the offset is a `u32`, so `offset=4294967296` — exactly 2^32 —
/// is the first illegal value.
///
/// ⛔ VALIDITY, NOT MALFORMEDNESS, AND THE SUITE SPLITS THEM. `offset=-1` is
/// malformed text ("unknown operator"); `offset=4294967296` parses perfectly
/// and is INVALID ("offset out of range"). A check in the wrong half makes the
/// other half's assertion pass for the wrong reason.
///
/// ⛔ memory64 RAISES THE CEILING, so the rule is skipped entirely when the
/// module declares an `i64` memory — the width is a property of the MEMORY,
/// not of the instruction, and applying the 32-bit bound there would reject
/// valid memory64 modules.
fn memarg_offset_range_reason(module: &Pair<Rule>) -> Option<String> {
    fn any_i64_memory(p: &Pair<Rule>) -> bool {
        if p.as_rule() == Rule::mem_type {
            return p.as_str().split_whitespace().next() == Some("i64");
        }
        p.clone().into_inner().any(|c| any_i64_memory(&c))
    }
    if any_i64_memory(module) {
        return None;
    }
    fn scan(p: &Pair<Rule>) -> Option<String> {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            for c in p.clone().into_inner() {
                // A `mem_arg` reaches an instruction wrapped in `instr_arg`;
                // accept it bare too rather than depend on which.
                let txt = match c.as_rule() {
                    Rule::mem_arg => c.as_str(),
                    Rule::instr_arg => match c.clone().into_inner().next() {
                        Some(i) if i.as_rule() == Rule::mem_arg => i.as_str(),
                        _ => continue,
                    },
                    _ => continue,
                };
                if let Some(v) = txt.trim().strip_prefix("offset=") {
                    if parse_wat_u128(v).is_some_and(|n| n > u32::MAX as u128) {
                        return Some("offset out of range".to_string());
                    }
                }
            }
        }
        for c in p.clone().into_inner() {
            if let Some(r) = scan(&c) {
                return Some(r);
            }
        }
        None
    }
    scan(module)
}

/// §3.3.4: a SIMD lane immediate must be in range for the instruction's shape.
/// `i8x16.extract_lane_s 16` names lane 16 of a 16-lane vector.
///
/// The bound comes from the SHAPE, not from the instruction family, so it is
/// read off the name: the `SHAPE.` prefix on extract/replace, and the access
/// WIDTH on `v128.loadN_lane`/`storeN_lane` (a `load8_lane` addresses the same
/// 16 lanes an `i8x16` does). `i8x16.shuffle` is the odd one — SIXTEEN
/// immediates, each selecting from the two input vectors' 32 lanes together.
///
/// ⛔ THE FIRST INTEGER IMMEDIATE, NOT THE FIRST IMMEDIATE. On the lane
/// load/stores the memarg comes first (`v128.load8_lane offset=0 1`), and on
/// the folded spellings the vector operand is a nested instruction rather than
/// an immediate. Filtering to integers is what makes one reading serve every
/// spelling — and picking the wrong token here would reject VALID modules,
/// which is the costly direction for a check that runs over every module.
fn lane_index_reason(module: &Pair<Rule>) -> Option<String> {
    /// Lanes addressable by this instruction, or `None` if it takes no lane.
    fn lanes_of(name: &str) -> Option<(u128, usize)> {
        // (bound, how many lane immediates)
        if name == "i8x16.shuffle" {
            return Some((32, 16));
        }
        if !name.ends_with("_lane") {
            return None;
        }
        let shape_lanes = |s: &str| -> Option<u128> {
            Some(match s {
                "i8x16" => 16,
                "i16x8" => 8,
                "i32x4" | "f32x4" => 4,
                "i64x2" | "f64x2" => 2,
                _ => return None,
            })
        };
        let (head, tail) = name.split_once('.')?;
        if head == "v128" {
            // load8_lane / store16_lane / … — the width names the lane count.
            let w = tail.trim_start_matches("load").trim_start_matches("store");
            let w = w.strip_suffix("_lane")?;
            return Some((
                match w {
                    "8" => 16,
                    "16" => 8,
                    "32" => 4,
                    "64" => 2,
                    _ => return None,
                },
                1,
            ));
        }
        // extract_lane_s / extract_lane_u / replace_lane
        Some((shape_lanes(head)?, 1))
    }

    fn scan(p: &Pair<Rule>) -> Option<String> {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            let name = head_keyword(p);
            if let Some((bound, count)) = lanes_of(&name) {
                let mut seen = 0usize;
                for c in p.clone().into_inner() {
                    let t = match c.as_rule() {
                        Rule::integer => c.as_str(),
                        Rule::instr_arg => match c.clone().into_inner().next() {
                            Some(i) if i.as_rule() == Rule::integer => i.as_str(),
                            _ => continue,
                        },
                        _ => continue,
                    };
                    if let Some(v) = parse_wat_u128(t.trim()) {
                        if v >= bound {
                            return Some("invalid lane index".to_string());
                        }
                        seen += 1;
                        if seen == count {
                            break;
                        }
                    }
                }
            }
        }
        for c in p.clone().into_inner() {
            if let Some(r) = scan(&c) {
                return Some(r);
            }
        }
        None
    }
    scan(module)
}

/// A type may reference only types declared BEFORE it, or types in its OWN
/// recursion group. Mutual recursion is legal exactly inside one `(rec …)`.
///
/// ```wast
/// (type $t1 (func (param (ref $t2))))   ;; forward, and NOT in $t2's group
/// (type $t2 (func (param (ref $t1))))
///   → "unknown type"
/// (rec (type (func (param (ref 1))))) (rec (type (func)))
///   → "unknown type"   ;; two SEPARATE groups, so still forward
/// ```
///
/// ⛔ A STANDALONE `(type …)` IS ITS OWN SINGLETON GROUP, so a self-reference
/// is legal and `r == idx` must NOT be flagged. Only a reference to a LATER
/// index in a DIFFERENT group is an error — which is why this needs the
/// rec-group boundary and could not be written before `rec_begin`/`rec_end`
/// reached the parse tree.
fn type_forward_reference_reason(module: &Pair<Rule>) -> Option<String> {
    let rec_of = rec_groups_of_module(module)?;
    let (types, names) = descriptor_type_table(module);
    let mut idx = 0usize;
    for field in module_fields(module) {
        if field.as_rule() != Rule::type_field {
            continue;
        }
        let mut refs: Vec<usize> = Vec::new();
        collect_concrete_type_refs(&field, &names, &mut refs);
        for r in refs {
            if r >= types.len() {
                return Some("unknown type".to_string());
            }
            if r > idx && rec_of.get(r) != rec_of.get(idx) {
                return Some("unknown type".to_string());
            }
        }
        idx += 1;
    }
    None
}

/// Every CONCRETE (`Heap::Concrete`) type index a type definition mentions.
/// Abstract heap types name no index and are skipped.
fn collect_concrete_type_refs(
    p: &Pair<Rule>,
    names: &HashMap<String, usize>,
    out: &mut Vec<usize>,
) {
    let t = p.as_str().trim();
    if t.starts_with("(ref") {
        if let Some(Vt::Ref(r)) = parse_vt(t, names) {
            if let Heap::Concrete(i) = r.heap {
                out.push(i);
            }
        }
    }
    for c in p.clone().into_inner() {
        collect_concrete_type_refs(&c, names, out);
    }
}

/// Source text on one line, runs of whitespace collapsed — fixture forms are
/// written across several lines and a diagnostic wants them inline.
fn squeeze_ws(s: &str) -> String {
    let mut out = String::new();
    let mut sp = false;
    for c in s.chars() {
        if c.is_whitespace() {
            sp = true;
            continue;
        }
        if sp && !out.is_empty() {
            out.push(' ');
        }
        sp = false;
        // The message is embedded in a string literal downstream.
        out.push(if c == '"' { '\'' } else { c });
    }
    out
}

/// `array.copy $dst $src` — the SOURCE element must be assignable to the
/// DESTINATION element.
///
/// ⛔ ITS OWN DIAGNOSTIC, NOT "type mismatch". The suite asserts "array types
/// do not match", and the operands are all perfectly well typed — this is a
/// relation between the two TYPE immediates, which is why the stack-typing
/// pass cannot see it.
fn array_copy_element_reason(module: &Pair<Rule>) -> Option<String> {
    let (types, names) = descriptor_type_table(module);
    fn scan(
        p: &Pair<Rule>,
        types: &[DescType],
        names: &HashMap<String, usize>,
    ) -> Option<String> {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr)
            && instr_head_name(p).as_deref() == Some("array.copy")
        {
            let idx: Vec<usize> = p
                .clone()
                .into_inner()
                .filter(|c| c.as_rule() == Rule::instr_arg)
                .map(|c| c.as_str().trim().to_string())
                .filter(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                .filter_map(|t| resolve_wast_index(&t, names))
                .collect();
            if let (Some(&d), Some(&sr)) = (idx.first(), idx.get(1)) {
                if let (Some(dt), Some(st)) = (types.get(d), types.get(sr)) {
                    if dt.kind == Some("array") && st.kind == Some("array") {
                        if let (Some(de), Some(se)) = (&dt.array_elem, &st.array_elem) {
                            if de != se {
                                match (parse_vt(se, names), parse_vt(de, names)) {
                                    (Some(x), Some(y)) if vt_subtype(&x, &y, types) => {}
                                    (Some(x), Some(y)) if provably_not_subtype(&x, &y, types) => {
                                        return Some("array types do not match".to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                }
            }
        }
        for c in p.clone().into_inner() {
            if let Some(r) = scan(&c, types, names) {
                return Some(r);
            }
        }
        None
    }
    scan(module, &types, &names)
}

/// Writing through an array whose element type is not `(mut …)`.
///
/// ⛔ MUTABILITY IS PER-ELEMENT AND OPT-IN. `(type $a (array i64))` declares an
/// IMMUTABLE array — `(array i64)` is not shorthand for `(array (mut i64))` —
/// so `array.set` on it is invalid however well-typed its operands are. Five
/// files in the GC suite assert this, and it is a property of the TYPE, which
/// is why the stack-typing pass cannot see it: every operand is correct.
///
/// `array.copy` names DEST first, then source; only the destination is
/// written. The read-only forms (`array.get`, `array.len`, `array.new*`) are
/// deliberately absent.
fn array_immutability_reason(module: &Pair<Rule>) -> Option<String> {
    let (types, names) = descriptor_type_table(module);
    fn scan(
        p: &Pair<Rule>,
        types: &[DescType],
        names: &HashMap<String, usize>,
    ) -> Option<String> {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            let head = instr_head_name(p).unwrap_or_default();
            if matches!(
                head.as_str(),
                "array.set" | "array.fill" | "array.copy" | "array.init_data" | "array.init_elem"
            ) {
                // The first index-shaped immediate is the array type written to.
                let ty = p
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .map(|c| c.as_str().trim().to_string())
                    .find(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                    .and_then(|t| resolve_wast_index(&t, names))
                    .and_then(|i| types.get(i));
                if let Some(t) = ty {
                    if t.kind == Some("array") && !t.array_elem_mut {
                        return Some("immutable array".to_string());
                    }
                }
            }
        }
        for c in p.clone().into_inner() {
            if let Some(r) = scan(&c, types, names) {
                return Some(r);
            }
        }
        None
    }
    scan(module, &types, &names)
}

/// §3.4.5 (exception handling): a tag's type must have NO results — a tag
/// describes the values an exception CARRIES, and throwing never produces
/// anything. `(tag (result i32))` is invalid, and so is the imported spelling
/// `(import "" "" (tag (result i32)))`.
///
/// ⛔ BOTH SPELLINGS OR NEITHER. The suite asserts the defined form and the
/// imported form with the same wording, and they reach the tree as different
/// rules — checking only `tag_field` answers one fixture of the two, which is
/// the same half-covered shape that left `br_on_null` behind.
fn tag_result_type_reason(module: &Pair<Rule>) -> Option<String> {
    fn has_result(p: &Pair<Rule>) -> bool {
        find_rule(p, Rule::tag_type).is_some_and(|tt| {
            tt.into_inner().any(|c| {
                c.as_rule() == Rule::result
                    // A bare `(result)` carries nothing and is fine; it is a
                    // result with a TYPE in it that makes the tag invalid.
                    && c.into_inner().next().is_some()
            })
        })
    }
    for field in module_fields(module) {
        let is_tag = match field.as_rule() {
            Rule::tag_field => true,
            Rule::import_field => find_rule(&field, Rule::import_desc)
                .is_some_and(|d| head_keyword(&d) == "tag"),
            _ => false,
        };
        if is_tag && has_result(&field) {
            return Some("non-empty tag result type".to_string());
        }
    }
    None
}

/// §6.6.4: a numeric `(type N)` must name a type that EXISTS — counting the
/// implicit types an inline signature defines.
///
/// ⛔ THE CENSUS DELIBERATELY OVER-COUNTS. `implicit_type_upper_bound` adds one
/// per inline signature with no dedup, which is the safe direction for a
/// name-resolution check — it can only miss, never over-fire. But missing is
/// exactly what happened: in `func.wast`'s fixture three functions carry inline
/// signatures of which only ONE is new, so the bound says 4 where the space
/// holds 2, and `(func (type 2))` looked in range.
///
/// `build_type_ctx` already computes the exact figure — its Pass 1b dedups
/// implicit types against both the explicit ones and each other, because the
/// INDICES depend on it. Reading that length here rather than re-deriving the
/// dedup keeps one answer in one place.
fn unknown_type_index_reason(module: &Pair<Rule>) -> Option<String> {
    let ctx = build_type_ctx(module);
    let n = ctx.type_sigs.len();
    fn walk(p: &Pair<Rule>, n: usize) -> Option<String> {
        if p.as_rule() == Rule::typeuse {
            if let Some(idx) = p.clone().into_inner().find(|c| c.as_rule() == Rule::index) {
                // Named indices are the census's question; this rule answers
                // only for the numeric spelling, whose bound it can prove.
                if let Ok(i) = idx.as_str().trim().parse::<usize>() {
                    if i >= n {
                        return Some(format!("unknown type {i}"));
                    }
                }
            }
        }
        for c in p.clone().into_inner() {
            if let Some(r) = walk(&c, n) {
                return Some(r);
            }
        }
        None
    }
    walk(module, n)
}

/// §3.4.4: a table's element type must be DEFAULTABLE unless the table is
/// written with an explicit initializer. Every slot starts life at the default
/// value, and a non-nullable reference has none — so `(table 0 (ref func))` is
/// invalid while `(table 0 funcref)` and `(table 1 (ref func) (ref.func $f))`
/// are both fine.
///
/// ⛔ THE SIZE DOES NOT MATTER, AND THE FIXTURES PIN THAT. `(table 0 (ref $f))`
/// is invalid even though it has no slots to fill: the rule is a property of
/// the TYPE, checked before any element count is considered. Reading it as
/// "only a non-empty table needs a default" discharges #16 and fails #17.
///
/// ⛔ THE `(elem …)` SPELLING IS AN INITIALIZER TOO. `table_field`'s second
/// branch — `(table $t funcref (elem $f …))` — supplies every slot, so it is
/// exempt; only the `table_type` branch with no trailing `folded_instr` can
/// trip this.
fn table_defaultable_reason(module: &Pair<Rule>) -> Option<String> {
    for field in module_fields(module) {
        if field.as_rule() != Rule::table_field {
            continue;
        }
        // The `(elem …)` branch has no `table_type` at all.
        let Some(tt) = find_rule(&field, Rule::table_type) else {
            continue;
        };
        // An IMPORTED table is initialised by whoever exports it.
        if find_rule(&field, Rule::import_inline).is_some() {
            continue;
        }
        if find_rule(&field, Rule::folded_instr).is_some() {
            continue;
        }
        let Some(rt) = find_rule(&tt, Rule::ref_val_type) else {
            continue;
        };
        let spelling = rt.as_str().trim();
        // Defaultable ⇔ nullable. The abbreviations (`funcref`, `externref`,
        // `anyref`, …) are all nullable by definition; only an explicit
        // `(ref …)` without `null` is not.
        let inner = spelling.trim_start_matches('(').trim_end_matches(')').trim();
        let non_null = inner
            .strip_prefix("ref")
            .is_some_and(|r| !r.trim_start().starts_with("null"));
        if non_null {
            return Some(format!("type mismatch: table element {spelling} is not defaultable"));
        }
    }
    None
}

/// §3.4.10: `ref.func x` is valid only when `x` is in `C.refs` — "the set of
/// function indices occurring in the module, EXCEPT in its functions or start
/// function". A function reachable only from inside a body, with nothing
/// declaring it, is invalid: `(module (func $f (drop (ref.func $f))))`.
///
/// ⛔ `start` DOES NOT DECLARE, and the suite pins that directly —
/// `(module (start $f) (func $f (drop (ref.func $f))))` is asserted invalid
/// even though `$f` plainly occurs in the module. Reading the rule as "occurs
/// anywhere outside a body" discharges the first fixture and fails this one.
///
/// ⛔ AN INLINE EXPORT DECLARES ITS OWN FUNCTION. `(func $f (export "a") …)`
/// desugars to a top-level export of `$f`, so the func field cannot simply be
/// skipped wholesale — that would reject a VALID module, which is the
/// expensive direction here: these rules run over every module the suite
/// compiles, not only the ones asserted invalid. Where this is unsure it
/// over-approximates the declared set, which can only cost a missed detection.
fn undeclared_func_ref_reason(module: &Pair<Rule>) -> Option<String> {
    fn collect_idents(p: &Pair<Rule>, out: &mut std::collections::HashSet<String>) {
        if matches!(p.as_rule(), Rule::index | Rule::id) {
            out.insert(p.as_str().trim().trim_start_matches('$').to_string());
        }
        for c in p.clone().into_inner() {
            collect_idents(&c, out);
        }
    }
    // Every `ref.func` use inside a function body, paired with nothing else —
    // the check is membership, so the use sites are just names.
    fn collect_uses(p: &Pair<Rule>, out: &mut Vec<String>) {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr)
            && head_keyword(p) == "ref.func"
        {
            if let Some(a) = p
                .clone()
                .into_inner()
                .flat_map(|c| {
                    if c.as_rule() == Rule::instr_arg {
                        c.into_inner().collect::<Vec<_>>()
                    } else {
                        vec![c]
                    }
                })
                .find(|c| matches!(c.as_rule(), Rule::index | Rule::id))
            {
                out.push(a.as_str().trim().trim_start_matches('$').to_string());
            }
        }
        for c in p.clone().into_inner() {
            collect_uses(&c, out);
        }
    }

    let mut declared: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut uses: Vec<String> = Vec::new();
    let mut func_index: usize = 0;
    for field in module_fields(module) {
        let inner = match field.as_rule() {
            Rule::module_field => match field.clone().into_inner().next() {
                Some(i) => i,
                None => continue,
            },
            _ => field.clone(),
        };
        match inner.as_rule() {
            // Bodies declare nothing; they only USE. The index the field
            // occupies is tracked either way, since a numeric `ref.func 0`
            // has to resolve against the same space.
            Rule::func_field => {
                let has_inline_export = inner
                    .clone()
                    .into_inner()
                    .any(|c| c.as_rule() == Rule::export_inline);
                if has_inline_export {
                    declared.insert(func_index.to_string());
                    if let Some(id) = census_id(&inner) {
                        declared.insert(id);
                    }
                }
                collect_uses(&inner, &mut uses);
                func_index += 1;
            }
            // The start function is explicitly outside `C.refs`.
            Rule::start_field => {}
            other => {
                if other == Rule::import_field && inner.as_str().contains("func") {
                    func_index += 1;
                }
                collect_idents(&inner, &mut declared);
            }
        }
    }
    uses.into_iter()
        .find(|u| !declared.contains(u))
        .map(|_| "undeclared function reference".to_string())
}

fn module_name_resolution_reason(module: &Pair<Rule>) -> Option<String> {
    let mut c = ModuleCensus::default();
    build_census(module, &mut c);
    name_resolution_walk(module.clone(), &c, &mut Vec::new(), 0)
}

fn name_resolution_walk(
    pair: Pair<Rule>,
    c: &ModuleCensus,
    locals: &mut Vec<std::collections::HashSet<String>>,
    depth: usize,
) -> Option<String> {
    let rule = pair.as_rule();
    // A function opens a LOCAL scope: its params and locals, by name. `local.get
    // $x` resolves there and nowhere else, so the scope is pushed for the body
    // and popped after — a module-level set would let one function see another's.
    let pushed = if rule == Rule::func_field {
        let mut names = std::collections::HashSet::new();
        for ch in pair.clone().into_inner() {
            match ch.as_rule() {
                Rule::typeuse => {
                    for pr in ch.into_inner() {
                        if pr.as_rule() == Rule::param {
                            if let Some(id) = census_id(&pr) {
                                names.insert(id);
                            }
                        }
                    }
                }
                Rule::local => {
                    if let Some(id) = census_id(&ch) {
                        names.insert(id);
                    }
                }
                _ => {}
            }
        }
        locals.push(names);
        true
    } else {
        false
    };

    // A `(type $t)` inside a typeuse names a declared type in every context it
    // appears — func signatures, `call_indirect`, GC ops.
    if rule == Rule::typeuse {
        if let Some(idx) = pair
            .clone()
            .into_inner()
            .find(|x| x.as_rule() == Rule::index)
        {
            if !index_resolves(&idx, &c.types) {
                if pushed {
                    locals.pop();
                }
                return Some("unknown type".to_string());
            }
        }
    }

    // ── unknown memory / unknown table ────────────────────────────────────
    //
    // ⛔ THE COMMON SHAPE IS "THE MODULE DECLARES NONE AT ALL", which needs no
    // memidx peel: `(module (func (drop (f32.load (i32.const 0)))))` is
    // "unknown memory" because memory 0 does not exist, not because the operand
    // was misread. That sidesteps the ambiguity that made `memory.init $d` and
    // `i32.load $m` unsafe to check by operand position — the count is asked,
    // never the index.
    //
    // An ACTIVE segment names a memory/table implicitly; a passive or
    // declarative one does not. `data_mode`/`elem_mode` present = active,
    // except `declare`.
    if c.memories.1 == 0 {
        let hits = match rule {
            Rule::data_field => pair
                .clone()
                .into_inner()
                .any(|x| x.as_rule() == Rule::data_mode),
            Rule::plain_instr | Rule::folded_instr => instr_head_name(&pair)
                .is_some_and(|n| instr_touches_memory(&n)),
            _ => false,
        };
        if hits {
            if pushed {
                locals.pop();
            }
            // ⛔ THE INDEX IS PART OF THE DIAGNOSTIC. The suite asserts
            // "unknown memory 0", and the comparison is one-directional — our
            // reason must CONTAIN the asserted text, so a bare "unknown
            // memory" does NOT discharge it. This arm only fires when the
            // module declares NO memory, so the reference is always memory 0;
            // naming it still satisfies the bare spelling too.
            return Some("unknown memory 0".to_string());
        }
    }
    if c.tables.1 == 0 {
        let hits = match rule {
            // ⛔ An elem's ELEMENT TYPE can be mistaken for its offset.
            // `elem_field` is `elem_mode? ~ ref_val_type?` and `elem_mode`'s
            // last alternative is a bare folded instruction, so `(elem $e
            // (ref 1))` — a PASSIVE segment declaring `(ref 1)` elements —
            // parses its type as an offset and looked active. An offset yields
            // an address; no `ref.*` form does, so the head decides.
            Rule::elem_field => pair.clone().into_inner().any(|x| {
                x.as_rule() == Rule::elem_mode
                    && x.as_str().trim() != "declare"
                    && !elem_mode_is_reference_type(&x)
            }),
            Rule::plain_instr | Rule::folded_instr => instr_head_name(&pair)
                .is_some_and(|n| instr_touches_table(&n)),
            _ => false,
        };
        if hits {
            if pushed {
                locals.pop();
            }
            // Same: this arm needs `tables == 0`, so the target is table 0.
            return Some("unknown table 0".to_string());
        }
    }

    // §3.4.6 (exception handling): `throw x` names a declared tag.
    // `(module (func (throw 0)))` has no tag section at all. The typing pass
    // never sees this — `throw` has no rule there and falls to the bail — so
    // it is answered off the census, exactly like `start`.
    if matches!(rule, Rule::plain_instr | Rule::folded_instr) && head_keyword(&pair) == "throw" {
        let idx = pair
            .clone()
            .into_inner()
            .flat_map(|x| {
                if x.as_rule() == Rule::instr_arg {
                    x.into_inner().collect::<Vec<_>>()
                } else {
                    vec![x]
                }
            })
            // ⛔ `Rule::integer`, NOT just `Rule::index` — and this file
            // already says so, at `peek_instr_tag_ref`: "an `instr_arg` spells
            // a bare tagidx as `integer`, not `index`; the `index` rule is
            // only reached where the grammar names it explicitly". Matching
            // `id`/`index` alone caught `throw $missing` and silently missed
            // `throw 0` in BOTH the folded and plain spellings — the same
            // sibling-spelling split, from re-deriving an immediate's rule set
            // instead of copying the helper that already got it right.
            .find(|x| matches!(x.as_rule(), Rule::index | Rule::id | Rule::integer));
        if let Some(idx) = idx {
            if !index_resolves(&idx, &c.tags) {
                if pushed {
                    locals.pop();
                }
                let t = idx.as_str().trim();
                return Some(match t.parse::<u64>() {
                    Ok(n) => format!("unknown tag {n}"),
                    Err(_) => "unknown tag".to_string(),
                });
            }
        }
    }

    // §3.4.9: the start function must exist. `(module (func) (start 1))` names
    // index 1 in a module with one function. Like an export descriptor this is
    // a reference OUTSIDE any instruction, so no instruction-level check could
    // ever reach it — and nothing else validated `start_field` at all.
    if rule == Rule::start_field {
        if let Some(idx) = pair
            .clone()
            .into_inner()
            .find(|x| x.as_rule() == Rule::index)
        {
            if !index_resolves(&idx, &c.funcs) {
                if pushed {
                    locals.pop();
                }
                return Some("unknown function".to_string());
            }
        }
    }

    // An EXPORT names an entity by index too — `(export "a" (func 0))` in a
    // module with no functions is "unknown function". Checked here because the
    // reference is in a descriptor, not an instruction.
    if rule == Rule::export_desc {
        let kind = pair
            .as_str()
            .trim_start_matches('(')
            .trim_start()
            .split_whitespace()
            .next()
            .unwrap_or("");
        if let Some(idx) = pair
            .clone()
            .into_inner()
            .find(|x| x.as_rule() == Rule::index)
        {
            let bad = match kind {
                "func" => (!index_resolves(&idx, &c.funcs)).then(|| "unknown function"),
                "global" => (!index_resolves(&idx, &c.globals)).then(|| "unknown global"),
                // Tables and memories are counted, not named, in the census —
                // a bare count is all an index needs.
                "table" => {
                    (!index_resolves(&idx, &c.tables)).then(|| "unknown table")
                }
                "memory" => (!index_resolves(&idx, &c.memories)).then(|| "unknown memory"),
                _ => None,
            };
            if let Some(msg) = bad {
                if pushed {
                    locals.pop();
                }
                // Name the index when it is written as one. The suite asserts
                // "unknown global 1" and "unknown data segment 1"; a `$name`
                // spelling has no number to quote, so it stays bare.
                let t = idx.as_str().trim();
                return Some(match t.parse::<u64>() {
                    Ok(n) => format!("{msg} {n}"),
                    Err(_) => msg.to_string(),
                });
            }
        }
    }
    // A `(ref $t)` VALUE TYPE names a declared type, and it can appear where no
    // typeuse does — `(type $x (func (param (ref $undeclared))))`. `heap_type`
    // is a single atomic token, so an `id`/`integer` there IS the reference;
    // the abstract spellings (`func`, `any`, `none`, …) are not.
    if rule == Rule::heap_type {
        let t = pair.as_str().trim();
        let concrete = t.strip_prefix('$').map(|n| (true, n.to_string())).or_else(|| {
            t.parse::<usize>().ok().map(|_| (false, t.to_string()))
        });
        if let Some((named, key)) = concrete {
            let ok = if named {
                c.types.0.contains(&key)
            } else {
                key.parse::<usize>().is_ok_and(|n| n < c.types.1)
            };
            if !ok {
                if pushed {
                    locals.pop();
                }
                return Some("unknown type".to_string());
            }
        }
    }
    if matches!(rule, Rule::plain_instr | Rule::folded_instr) {
        let mut name: Option<String> = None;
        let mut first_idx: Option<Pair<Rule>> = None;
        for ch in pair.clone().into_inner() {
            match ch.as_rule() {
                Rule::instr_name if name.is_none() => name = Some(ch.as_str().to_string()),
                Rule::instr_arg if first_idx.is_none() => {
                    if let Some(i) = ch.clone().into_inner().next() {
                        if matches!(i.as_rule(), Rule::index | Rule::id | Rule::integer) {
                            first_idx = Some(if i.as_rule() == Rule::index { i } else { ch });
                        }
                    }
                }
                _ => {}
            }
        }
        // How many index-shaped immediates this instruction carries. ⛔ THE
        // MEMIDX PEEL DECIDES WHAT OPERAND 0 MEANS: `memory.init d` names a
        // DATA segment, `memory.init m d` names a memory THEN a segment. With
        // one immediate the position is certain; with two it is not, and
        // guessing reports the wrong "unknown X" — which the message
        // comparison turns into a visible failure rather than a silent pass.
        let idx_count = pair
            .clone()
            .into_inner()
            .filter(|ch| ch.as_rule() == Rule::instr_arg)
            .filter(|ch| {
                ch.clone()
                    .into_inner()
                    .next()
                    .is_some_and(|i| matches!(i.as_rule(), Rule::index | Rule::id | Rule::integer))
            })
            .count();
        if let (Some(n), Some(idx)) = (name.as_deref(), first_idx) {
            let base = n.split_once("@@").map(|(b, _)| b).unwrap_or(n);
            let bad = match base {
                // Operand 0 is unambiguously that entity for these.
                "global.get" | "global.set" => {
                    (!index_resolves(&idx, &c.globals)).then(|| "unknown global")
                }
                "call" | "ref.func" | "return_call" => {
                    (!index_resolves(&idx, &c.funcs)).then(|| "unknown function")
                }
                "data.drop" => {
                    (!index_resolves(&idx, &c.data_segs)).then(|| "unknown data segment")
                }
                // Single immediate ⇒ it IS the segment index (see the peel note).
                "memory.init" if idx_count == 1 => {
                    (!index_resolves(&idx, &c.data_segs)).then(|| "unknown data segment")
                }
                "elem.drop" => {
                    (!index_resolves(&idx, &c.elem_segs)).then(|| "unknown elem segment")
                }
                "table.init" if idx_count == 1 => {
                    (!index_resolves(&idx, &c.elem_segs)).then(|| "unknown elem segment")
                }
                // A bare `ref` is not an instruction — it is `(ref $t)` / `(ref N)`,
                // a reference TYPE the grammar folded here (see
                // `elem_mode_is_reference_type`). Its index names a declared type.
                "ref" => (!index_resolves(&idx, &c.types)).then(|| "unknown type"),
                "local.get" | "local.set" | "local.tee" => {
                    let t = idx.as_str().trim();
                    match t.strip_prefix('$') {
                        Some(nm) => (!locals.last().is_some_and(|sc| sc.contains(nm)))
                            .then(|| "unknown local"),
                        // A NUMERIC local index needs the param+local count,
                        // which the typeuse may state by reference to a type —
                        // not resolvable from this scan alone.
                        None => None,
                    }
                }
                _ => None,
            };
            if let Some(msg) = bad {
                if pushed {
                    locals.pop();
                }
                // Name the index when it is written as one. The suite asserts
                // "unknown global 1" and "unknown data segment 1"; a `$name`
                // spelling has no number to quote, so it stays bare.
                let t = idx.as_str().trim();
                return Some(match t.parse::<u64>() {
                    Ok(n) => format!("{msg} {n}"),
                    Err(_) => msg.to_string(),
                });
            }
        }
    }

    // ── unknown label ─────────────────────────────────────────────────────
    // `br N` targets the Nth enclosing label, 0 = innermost, and the FUNCTION
    // BODY is itself a label — so inside a function with no nested blocks the
    // only legal depth is 0, and `(func (br 1))` is out of range.
    // `br_table` names several at once and every one of them must be in range.
    //
    // ⛔ Folded spelling only. In the plain spelling `block … end` is a FLAT
    // token sequence, not a nested pair, so the tree carries no nesting to
    // count; guessing there would over-flag. Every fixture in the suite writes
    // these folded.
    //
    // ⛔ `instr_head_name` CANNOT SEE A FOLDED BLOCK. In `"(" ~ "block" ~ id? ~
    // block_type* ~ instr*` the keyword is a grammar LITERAL, so the pair has
    // no `instr_name` child and the lookup returns None — the depth counter
    // never incremented for the exact construct it exists to count. Every
    // nested `br 1` therefore read as out of range: 28 fixtures in block.wast
    // alone were reported "unknown label" when they assert "type mismatch".
    // The keyword has to be read off the TEXT, which is what `head_keyword`
    // does for the same reason in the flattener.
    let inner_depth = if matches!(rule, Rule::folded_instr)
        && matches!(head_keyword(&pair).as_str(), "block" | "loop" | "if")
    {
        depth + 1
    } else {
        depth
    };
    if matches!(rule, Rule::plain_instr | Rule::folded_instr) {
        if let Some(n) = instr_head_name(&pair) {
            if matches!(n.as_str(), "br" | "br_if" | "br_table") {
                for arg in pair.clone().into_inner() {
                    if arg.as_rule() != Rule::instr_arg {
                        continue;
                    }
                    let Some(i) = arg.clone().into_inner().next() else {
                        continue;
                    };
                    if i.as_rule() != Rule::integer {
                        continue;
                    }
                    if let Some(v) = parse_wat_u128(i.as_str()) {
                        if v > depth as u128 {
                            if pushed {
                                locals.pop();
                            }
                            return Some("unknown label".to_string());
                        }
                    }
                }
            }
        }
    }

    for child in pair.into_inner() {
        if let Some(r) = name_resolution_walk(child, c, locals, inner_depth) {
            if pushed {
                locals.pop();
            }
            return Some(r);
        }
    }
    if pushed {
        locals.pop();
    }
    None
}

fn module_invalid_walk(
    pair: Pair<Rule>,
    export_names: &mut std::collections::HashSet<String>,
) -> Option<String> {
    // Custom Descriptors' type-section rules. Whole-table, so they run ONCE
    // per module rather than per pair — this recursion reaches `Rule::module`
    // exactly once on its way down.
    if pair.as_rule() == Rule::module {
        if let Some(r) = descriptor_invalid_reason(&pair) {
            return Some(r);
        }
    }
    // A module's export names must be pairwise distinct — over ALL of
    // funcs/tables/memories/globals/tags, which is why one set covers the
    // inline and the standalone spellings together.
    if matches!(pair.as_rule(), Rule::export_inline | Rule::export_field) {
        if let Some(sname) = pair
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::string)
        {
            let n = unquote(sname.as_str());
            if !export_names.insert(n) {
                return Some("duplicate export name".to_string());
            }
        }
    }
    // Name resolution is a WHOLE-MODULE question — it needs the census — so it
    // runs once when the walk reaches the module, not per pair.
    if pair.as_rule() == Rule::module {
        if let Some(r) = module_name_resolution_reason(&pair) {
            return Some(r);
        }
        // Kept alongside the index-space check in the stack-typing pass: this
        // one still answers when that pass BAILS on a module it cannot fully
        // type, and both give the same wording.
        if let Some(r) = immutable_global_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = undeclared_func_ref_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = lane_index_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = unknown_type_index_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = table_defaultable_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = tag_result_type_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = memarg_offset_range_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = module_unknown_local_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = type_forward_reference_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = array_immutability_reason(&pair) {
            return Some(r);
        }
        if let Some(r) = array_copy_element_reason(&pair) {
            return Some(r);
        }
    }
    // ── limits ────────────────────────────────────────────────────────────
    // `mem_type`/`table_type` are `addr_type? ~ integer ~ integer?`, so the
    // limits are the one or two integers after an optional `i32`/`i64`.
    //
    // ⛔ Three distinct diagnostics, and they are not interchangeable:
    // a memory over 65536 PAGES is "memory size", a table over 2^32 ENTRIES is
    // "table size", and min > max is "size minimum must not be greater than
    // maximum" for both. The suite asserts each by its own wording.
    if matches!(pair.as_rule(), Rule::mem_type | Rule::table_type) {
        let is_mem = pair.as_rule() == Rule::mem_type;
        // memory64 raises the page ceiling; a 64-bit table's is 2^64.
        let is64 = pair.as_str().split_whitespace().next() == Some("i64");
        let nums: Vec<u128> = pair
            .clone()
            .into_inner()
            .filter(|c| c.as_rule() == Rule::integer)
            .filter_map(|c| parse_wat_u128(c.as_str()))
            .collect();
        // A limit too large to be a u128 is malformed text, not invalid, and
        // `parse_wat_u128` declining leaves it to that half.
        let ceiling: u128 = match (is_mem, is64) {
            (true, false) => 65536,               // 4 GiB in 64 KiB pages
            (true, true) => 1u128 << 48,          // memory64 page ceiling
            // ⛔ INCLUSIVE. A table holds at most 2^32 - 1 entries, so
            // `0x1_0000_0000` is the FIRST illegal value — `> 1<<32` let it
            // through. A memory's 65536 pages, by contrast, IS legal.
            (false, false) => (1u128 << 32) - 1,  // table entries
            (false, true) => u128::MAX,
        };
        if let Some(&min) = nums.first() {
            if min > ceiling {
                return Some(if is_mem { "memory size" } else { "table size" }.to_string());
            }
            if let Some(&max) = nums.get(1) {
                if max > ceiling {
                    return Some(if is_mem { "memory size" } else { "table size" }.to_string());
                }
                if min > max {
                    return Some("size minimum must not be greater than maximum".to_string());
                }
            }
        }
    }
    // ── constant expression required ──────────────────────────────────────
    // A global initializer and an active segment's offset are CONSTANT
    // EXPRESSIONS: only the const forms, `global.get`, the GC allocations, and
    // the extended-const arithmetic. `(global f32 (f32.neg (f32.const 0)))`
    // and `(data (nop))` are the shapes the suite asserts.
    if matches!(
        pair.as_rule(),
        Rule::global_field | Rule::data_mode | Rule::elem_mode | Rule::elem_item
    ) {
        // An IMPORTED global has no initializer to check.
        let imported = pair.as_rule() == Rule::global_field
            && pair
                .clone()
                .into_inner()
                .any(|c| c.as_rule() == Rule::import_inline);
        // ⛔ SAME GRAMMAR AMBIGUITY, THIRD SYMPTOM. `(elem $e (ref 1))`'s
        // element TYPE matches `elem_mode`'s bare-folded-instr alternative, so
        // it looks like an offset here too — and `ref` is not a constant
        // expression, so this reported "constant expression required" for a
        // segment that has no offset at all. It is a type; skip it.
        let is_ref_type = pair.as_rule() == Rule::elem_mode
            && elem_mode_is_reference_type(&pair);
        if !imported && !is_ref_type {
            if let Some(bad) = first_non_const_instr(&pair) {
                let _ = bad;
                return Some("constant expression required".to_string());
            }
        }
    }
    if matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
        let mut name: Option<String> = None;
        let mut align: Option<u32> = None;
        let mut first_int: Option<i64> = None;
        for c in pair.clone().into_inner() {
            match c.as_rule() {
                Rule::instr_name if name.is_none() => name = Some(c.as_str().to_string()),
                Rule::instr_arg => {
                    let t = c.as_str().trim();
                    if let Some(d) = t.strip_prefix("align=") {
                        align = d.parse::<u32>().ok();
                    } else if first_int.is_none() {
                        if let Some(inner) = c.clone().into_inner().next() {
                            if inner.as_rule() == Rule::integer {
                                first_int = inner.as_str().parse::<i64>().ok();
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        if let Some(n) = name.as_deref() {
            let base = n.split_once("@@").map(|(b, _)| b).unwrap_or(n);
            // `align` states the access alignment in BYTES and must not exceed
            // the width the instruction actually touches. `natural_align_bytes`
            // is the opcode table's own answer — not a list kept here.
            if let (Some(a), Some(op)) = (align, vybe_runtime::opcode::Op::from_wasm_name(base)) {
                if let Some(natural) = op.natural_align_bytes() {
                    if a > natural {
                        return Some("alignment must not be larger than natural".to_string());
                    }
                }
            }
            // A lane immediate must be inside the vector's lane count, and the
            // MNEMONIC states that count: `i8x16.*` has 16, `v128.load32_lane`
            // has 128/32 = 4.
            if let (Some(lanes), Some(idx)) = (mnemonic_lane_count(base), first_int) {
                if idx < 0 || idx >= lanes as i64 {
                    return Some("invalid lane index".to_string());
                }
            }
        }
    }
    for child in pair.into_inner() {
        if let Some(r) = module_invalid_walk(child, export_names) {
            return Some(r);
        }
    }
    None
}

// ── Custom Descriptors: type-section validity ────────────────────────────────
//
// The proposal's structural rules — the ones that need no stack-typing pass.
// Kept OUT of `module_invalid_walk`'s inline arms deliberately: those are the
// rules that apply to every module (alignment, lane index, duplicate export)
// and are checked per-pair during a generic recursion, while these need the
// whole type table and its rec-group structure at once.
//
// ⛔ RETURNS `None` FOR ANYTHING IT DOES NOT RECOGNISE. Never a generic
// "invalid module": `assert_invalid` compares the expected diagnostic, so a
// rule firing with the wrong message discharges an assertion it was never
// written for — the same lie as discharging without looking, one level down.

/// One type's descriptor-related declaration, as written.
#[derive(Default, Clone)]
struct DescType {
    /// Which recursion group this type belongs to. Every standalone `(type …)`
    /// is its own singleton group; a `(rec …)` shares one.
    rec_group: usize,
    is_struct: bool,
    /// `(descriptor N)` — N is the type that DESCRIBES this one.
    descriptor: Option<usize>,
    /// `(describes N)` — N is the type this one is the descriptor FOR.
    describes: Option<usize>,
    /// Declared supertypes: the `index*` of `(sub $parent … )`. Empty for a
    /// `(sub …)` with no parent and for a type with no `sub` at all.
    supers: Vec<usize>,
    /// Declared struct field count — the operand count an allocation needs
    /// before its descriptor.
    fields: usize,
    /// The declared field type SPELLINGS, in order, for subtype comparison.
    field_types: Vec<String>,
    /// Per field, its `$name` when one was written — PARALLEL to `field_types`.
    ///
    /// ⛔ FIELD NAMES ARE SCOPED TO THEIR STRUCT, NOT THE MODULE. Two types may
    /// each declare `$x`, and `struct.get 0 $x` means type 0's. A module-wide
    /// map would answer with whichever was seen last, which is a confident
    /// wrong answer where the fixtures differ by exactly that.
    field_names: Vec<Option<String>>,
    /// ISO-RECURSIVE CANONICAL ID. Two declarations with this id equal ARE the
    /// same type, even in different rec groups. `None` when the module's
    /// groups could not be canonicalised, which leaves every comparison
    /// undecided rather than wrong.
    canon: Option<usize>,
    /// Was this type written with an explicit `(sub …)`?
    ///
    /// ⛔ A PLAIN `(type $t (struct))` IS FINAL. It abbreviates
    /// `(sub final (struct))`, so nothing may declare it as a supertype —
    /// which is what `(type $s (sub $t (struct)))` asserts "sub type" for.
    /// Openness is opt-in, not the default.
    has_sub: bool,
    /// `(sub final …)` — explicitly closed.
    is_final: bool,
    /// `struct` / `array` / `func`. A subtype must have the SAME form as its
    /// supertype; a struct may not extend a func.
    kind: Option<&'static str>,
    /// A func type's param and result SPELLINGS. Needed because function
    /// subtyping compares SIGNATURES, which no other field records.
    func_sig: Option<(Vec<String>, Vec<String>)>,
    /// For an `array` type: is its ELEMENT declared `(mut …)`?
    ///
    /// ⛔ NOT DERIVABLE FROM `field_types`. `array_type` is
    /// `"(" ~ "array" ~ (field_def | storage_type) ~ ")"`, and the common
    /// spelling `(array (mut i64))` is the bare `storage_type` branch — no
    /// `field_def` at all — so the field collector sees nothing and "no
    /// element" is indistinguishable from "immutable element".
    array_elem_mut: bool,
    /// For an `array` type: its ELEMENT type spelling, `mut` stripped.
    array_elem: Option<String>,
}

/// The rec-group id of every type in declaration order.
///
/// ⛔ UNAVAILABLE TODAY, AND THAT IS A GRAMMAR FACT, NOT AN OVERSIGHT.
/// `rec_group` is a SILENT pest rule (`_{ … }`), so a `(rec …)`'s members are
/// spliced straight into the module and the boundary never reaches the parse
/// tree. The grammar comment justifies that: "a TYPE-IDENTITY property, not a
/// structural one". Validation is the consumer that disproves it — these two
/// modules differ ONLY by the boundary and the spec gives them DIFFERENT
/// diagnostics:
///
/// ```wast
/// (type (descriptor 1) (struct)) (type (struct))
///   → "descriptor type is outside rec group"
/// (rec (type (descriptor 1) (struct)) (type (struct)))
///   → "type is not described by its descriptor"
/// ```
///
/// So the boundary decides WHICH message is right, and without it every rule
/// below is unsafe — not just the two that name rec groups. Returning `None`
/// here disables the whole helper, which is the honest state: those assertions
/// fail as "the module validated" rather than passing for a made-up reason.
///
/// Making `rec_group` non-silent is not the fix: `Rule::module_field` is
/// matched at 17 one-level sites in this file and every one would need a
/// flatten. A zero-width boundary marker inside the silent rule would be
/// invisible to all 17 (they filter FOR `module_field`) and is what this
/// function is waiting on.
fn rec_groups_of_module(module: &Pair<Rule>) -> Option<Vec<usize>> {
    // `rec_group` is silent, so a `(rec …)`'s members stay spliced into the
    // module exactly as before — but it now emits `rec_begin` / `rec_end`
    // markers around them, and those ARE in the tree. Walking the module's
    // children in order therefore recovers the boundary without any consumer
    // seeing a wrapper: all 17 `Rule::module_field` loops filter for
    // module_field and skip the markers.
    let mut groups: Vec<usize> = Vec::new();
    let mut next_group = 0usize;
    let mut current_rec: Option<usize> = None;
    for child in module.clone().into_inner() {
        match child.as_rule() {
            Rule::rec_begin => {
                current_rec = Some(next_group);
                next_group += 1;
            }
            Rule::rec_end => current_rec = None,
            Rule::module_field => {
                let is_type = child
                    .clone()
                    .into_inner()
                    .next()
                    .is_some_and(|t| t.as_rule() == Rule::type_field);
                if is_type {
                    // Inside a `(rec …)` every member shares one group;
                    // outside it, each `(type …)` is its own singleton.
                    groups.push(current_rec.unwrap_or_else(|| {
                        let g = next_group;
                        next_group += 1;
                        g
                    }));
                }
            }
            _ => {}
        }
    }
    Some(groups)
}

/// Resolve a `(descriptor …)` / `(describes …)` operand — a numeric index or a
/// `$name` — against the module's declared type names.
fn desc_clause_target(clause: &Pair<Rule>, names: &HashMap<String, usize>) -> Option<usize> {
    let idx = clause
        .clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::index)?;
    let text = idx.as_str().trim();
    match text.parse::<usize>() {
        Ok(n) => Some(n),
        // A `$name` that resolves to nothing is left unresolved rather than
        // guessed at: "unknown type" is its own diagnostic and belongs to
        // whoever implements it, not to a silent fallback here.
        Err(_) => names.get(text).copied(),
    }
}

/// Read the descriptor declarations off a module's type section, in
/// declaration order.
fn descriptor_type_table(module: &Pair<Rule>) -> (Vec<DescType>, HashMap<String, usize>) {
    // Pass 1: `$name` → index, so a clause may name a type declared later.
    let mut names: HashMap<String, usize> = HashMap::new();
    let mut type_fields: Vec<Pair<Rule>> = Vec::new();
    for child in module.clone().into_inner() {
        if child.as_rule() != Rule::module_field {
            continue;
        }
        if let Some(tf) = child.into_inner().next() {
            if tf.as_rule() == Rule::type_field {
                if let Some(id) = tf.clone().into_inner().find(|c| c.as_rule() == Rule::id) {
                    names.insert(id.as_str().trim().to_string(), type_fields.len());
                }
                type_fields.push(tf);
            }
        }
    }
    // Pass 2: the clauses. `desc_clauses` is silent, so the two clause pairs
    // are spliced either directly into `type_field` or into its `sub_type` —
    // `(type (descriptor $d) (struct))` vs `(type (sub (descriptor $d) …))`.
    // Searching both levels is what makes the `sub` spellings work; the
    // proposal's own fixtures use `sub`, `sub final` and `sub $parent` forms.
    let table: Vec<DescType> = type_fields
        .iter()
        .map(|tf| {
            let mut t = DescType::default();
            let mut scan = |p: &Pair<Rule>| {
                for c in p.clone().into_inner() {
                    match c.as_rule() {
                        Rule::describes_clause => {
                            t.describes = desc_clause_target(&c, &names);
                        }
                        Rule::descriptor_clause => {
                            t.descriptor = desc_clause_target(&c, &names);
                        }
                        Rule::composite_type => {
                            t.field_types = c
                                .clone()
                                .into_inner()
                                .next()
                                .map(|k| {
                                    k.into_inner()
                                        .filter(|f| f.as_rule() == Rule::field_def)
                                        // ⛔ ONE `field_def` CAN DECLARE SEVERAL
                                        // FIELDS. The grammar is
                                        // `"(" ~ "field" ~ (id ~ storage_type
                                        //  | storage_type*) ~ ")"`, so
                                        // `(field i32 i32)` — which the GC
                                        // suite uses — is TWO fields in one
                                        // node. Counting nodes undercounts,
                                        // which both shrinks the allocation
                                        // arity and silently shortens the
                                        // sub/super field comparison (a `zip`
                                        // stops at the shorter side, so the
                                        // extra fields go unchecked).
                                        .flat_map(|f| {
                                            f.into_inner()
                                                .filter(|x| x.as_rule() != Rule::id)
                                                .map(|x| x.as_str().trim().to_string())
                                                .collect::<Vec<_>>()
                                        })
                                        .collect::<Vec<String>>()
                                })
                                .unwrap_or_default();
                            // The names, kept PARALLEL to `field_types`. A
                            // `field_def` that carries an id declares exactly
                            // one field (`id ~ storage_type`); the multi-field
                            // spelling `(field i32 i32)` carries none, so each
                            // of its storage types contributes a `None`.
                            t.field_names = c
                                .clone()
                                .into_inner()
                                .next()
                                .map(|k| {
                                    k.into_inner()
                                        .filter(|f| f.as_rule() == Rule::field_def)
                                        .flat_map(|f| {
                                            let id = f
                                                .clone()
                                                .into_inner()
                                                .find(|x| x.as_rule() == Rule::id)
                                                .map(|x| x.as_str().to_string());
                                            let n = f
                                                .into_inner()
                                                .filter(|x| x.as_rule() != Rule::id)
                                                .count();
                                            (0..n)
                                                .map(|i| if i == 0 { id.clone() } else { None })
                                                .collect::<Vec<_>>()
                                        })
                                        .collect::<Vec<Option<String>>>()
                                })
                                .unwrap_or_default();
                            t.fields = t.field_types.len();
                            t.is_struct = c
                                .clone()
                                .into_inner()
                                .next()
                                .is_some_and(|k| {
                                    matches!(k.as_rule(), Rule::struct_type | Rule::struct_subtype)
                                });
                            // ⛔ THE FORM IS A LITERAL KEYWORD, so it is not a
                            // pair — read it off the text, the same way the
                            // folded block keyword has to be.
                            let head = c.as_str().trim_start_matches('(').trim_start();
                            t.kind = match head.split(|ch: char| ch.is_whitespace() || ch == '(' || ch == ')').next() {
                                Some("struct") => Some("struct"),
                                Some("array") => Some("array"),
                                Some("func") => Some("func"),
                                _ => t.kind,
                            };
                            if t.kind == Some("array") {
                                // Read it off the `array_type` text: the
                                // element is `(mut …)` or it is not.
                                if let Some(at) = find_rule(&c, Rule::array_type) {
                                    let inner = at
                                        .as_str()
                                        .trim()
                                        .trim_start_matches('(')
                                        .trim_start()
                                        .strip_prefix("array")
                                        .unwrap_or("")
                                        .trim_start();
                                    let inner = inner
                                        .strip_prefix("(field")
                                        .map(|r| r.trim_start())
                                        .unwrap_or(inner);
                                    t.array_elem_mut = inner.starts_with("(mut");
                                    // Strip the `mut` wrapper so the element
                                    // can be compared as an ordinary type.
                                    let e = if t.array_elem_mut {
                                        inner
                                            .strip_prefix("(mut")
                                            .map(|r| r.trim().trim_end_matches(')').trim())
                                            .unwrap_or(inner)
                                    } else {
                                        inner
                                    };
                                    let e = e.trim().trim_end_matches(')').trim();
                                    if !e.is_empty() {
                                        t.array_elem = Some(e.to_string());
                                    }
                                }
                            }
                            if t.kind == Some("func") {
                                let mut ps = Vec::new();
                                let mut rs = Vec::new();
                                collect_params_results(&c, &mut ps, &mut rs);
                                t.func_sig = Some((ps, rs));
                            }
                        }
                        _ => {}
                    }
                }
            };
            scan(tf);
            if let Some(sub) = tf
                .clone()
                .into_inner()
                .find(|c| c.as_rule() == Rule::sub_type)
            {
                t.has_sub = true;
                // `final` is a bare literal in the grammar, so it never
                // reaches the tree as a pair.
                t.is_final = sub
                    .as_str()
                    .trim_start_matches('(')
                    .trim_start()
                    .strip_prefix("sub")
                    .map(|r| r.trim_start().starts_with("final"))
                    .unwrap_or(false);
                scan(&sub);
                // `sub_type = { "(" ~ "sub" ~ "final"? ~ index* ~ desc_clauses
                // ~ composite_type ~ ")" }` — the supertypes are the DIRECT
                // `index` children. The clauses' own indices are nested inside
                // `describes_clause` / `descriptor_clause`, so they are not
                // picked up here.
                for c in sub.into_inner() {
                    if c.as_rule() == Rule::index {
                        let text = c.as_str().trim();
                        if let Some(n) = text
                            .parse::<usize>()
                            .ok()
                            .or_else(|| names.get(text).copied())
                        {
                            t.supers.push(n);
                        }
                    }
                }
            }
            t
        })
        .collect();
    let mut table = table;
    if let Some(ids) = canonical_type_ids(module, &names) {
        for (t, id) in table.iter_mut().zip(ids) {
            t.canon = Some(id);
        }
    }
    (table, names)
}

/// The ISO-RECURSIVE canonical id of every type, in declaration order.
///
/// ⛔ TWO DECLARATIONS CAN BE THE SAME TYPE. WASM identifies types by the
/// STRUCTURE of their whole recursion group, not by where they were written —
/// so `(rec (type $f1 (sub (func))) …)` and an identically shaped
/// `(rec (type $f2 (sub (func))) …)` make `(ref $f1)` and `(ref $f2)`
/// interchangeable. `gc/type-subtyping.wast` is built to test exactly that,
/// and without this the declared-`sub`-chain walk answers "not a subtype" for
/// a perfectly valid module.
///
/// Computable in one forward pass ONLY because a type may reference just
/// earlier types or its own group — the rule `type_forward_reference_reason`
/// enforces. Each group is normalised with intra-group references written
/// group-relative (`#n`) and outward ones by the referent's ALREADY canonical
/// group (`@g.n`), then interned: identical normal form ⇒ identical group.
fn canonical_type_ids(module: &Pair<Rule>, names: &HashMap<String, usize>) -> Option<Vec<usize>> {
    let rec_of = rec_groups_of_module(module)?;
    let fields: Vec<Pair<Rule>> = module_fields(module)
        .into_iter()
        .filter(|f| f.as_rule() == Rule::type_field)
        .collect();
    if fields.len() != rec_of.len() {
        return None;
    }
    let ngroups = rec_of.iter().copied().max().map_or(0, |m| m + 1);
    let mut members: Vec<Vec<usize>> = vec![Vec::new(); ngroups];
    for (i, &g) in rec_of.iter().enumerate() {
        members[g].push(i);
    }
    let mut group_canon: Vec<Option<usize>> = vec![None; ngroups];
    let mut group_interned: HashMap<String, usize> = HashMap::new();
    let mut type_interned: HashMap<(usize, usize), usize> = HashMap::new();
    let mut canon = vec![0usize; fields.len()];
    let mut order: Vec<usize> = (0..ngroups).collect();
    order.sort_by_key(|&g| members[g].first().copied().unwrap_or(usize::MAX));
    for g in order {
        if members[g].is_empty() {
            continue;
        }
        let start = members[g][0];
        let mut form = String::new();
        for &i in &members[g] {
            form.push_str(&canon_form(
                &fields[i], names, &rec_of, g, start, &group_canon, &members,
            )?);
            form.push('|');
        }
        let next = group_interned.len();
        let gid = *group_interned.entry(form).or_insert(next);
        group_canon[g] = Some(gid);
        for (off, &i) in members[g].iter().enumerate() {
            let n = type_interned.len();
            canon[i] = *type_interned.entry((gid, off)).or_insert(n);
        }
    }
    Some(canon)
}

/// One type's structure as a normal-form string.
///
/// ⛔ TEXT, NOT A PAIR WALK. Pest does not capture string literals, so `"ref"`
/// and `"null"` are absent from the tree — a structural walk cannot tell
/// `(ref $t)` from `(ref null $t)` and would call two different types equal.
/// The type's OWN id is dropped (a name is not part of its structure); every
/// other `$name` or bare integer inside a type definition IS a type index.
fn canon_form(
    field: &Pair<Rule>,
    names: &HashMap<String, usize>,
    rec_of: &[usize],
    group: usize,
    start: usize,
    group_canon: &[Option<usize>],
    members: &[Vec<usize>],
) -> Option<String> {
    let src = field.as_str();
    let bytes = src.as_bytes();
    let mut out = String::new();
    let mut i = 0usize;
    let mut seen_type_kw = false;
    let mut expect_own_id = false;
    while i < bytes.len() {
        let c = bytes[i] as char;
        if c == ';' && src[i..].starts_with(";;") {
            i = src[i..].find('\n').map_or(bytes.len(), |k| i + k);
            continue;
        }
        if c.is_whitespace() {
            if !out.ends_with(' ') {
                out.push(' ');
            }
            i += 1;
            continue;
        }
        if c == '(' || c == ')' {
            out.push(c);
            i += 1;
            continue;
        }
        let mut end = bytes.len();
        for (j, c2) in src[i..].char_indices() {
            if c2.is_whitespace() || c2 == '(' || c2 == ')' {
                end = i + j;
                break;
            }
        }
        let tok = &src[i..end];
        i = end;
        if tok == "type" && !seen_type_kw {
            seen_type_kw = true;
            expect_own_id = true;
            out.push_str(tok);
            continue;
        }
        if expect_own_id {
            expect_own_id = false;
            if tok.starts_with('$') {
                continue;
            }
        }
        let idx = match tok.strip_prefix('$') {
            Some(bare) => names.get(tok).or_else(|| names.get(bare)).copied(),
            None => tok.parse::<usize>().ok(),
        };
        match idx {
            Some(r) if r < rec_of.len() => {
                if rec_of[r] == group {
                    out.push_str(&format!("#{}", r - start));
                } else {
                    let gid = group_canon.get(rec_of[r]).copied().flatten()?;
                    let off = members[rec_of[r]].iter().position(|&x| x == r)?;
                    out.push_str(&format!("@{gid}.{off}"));
                }
            }
            _ => out.push_str(tok),
        }
    }
    Some(out)
}

/// The Custom Descriptors type-section rules.
///
/// ⚠ THE ORDER OF THESE CHECKS IS SPEC, NOT TASTE. Two modules can violate
/// several rules at once and the suite pins exactly one message for each, so a
/// reordering produces a WRONG diagnostic rather than a differently-worded
/// right one. Derived from the fixtures, verified case by case:
///
///   1. rec-group membership  — `(descriptor D)` then `(describes S)`
///   2. struct-ness of the CLAUSE CARRIER (not of its target: in
///      `(rec (type (descriptor 1) (func)) (type (describes 0) (struct)))`
///      the descriptor IS a struct and the message is still "descriptor type
///      must be a struct", naming the clause the offending type carries)
///   3. forward use of a described type
///   4. pair agreement in both directions
fn descriptor_invalid_reason(module: &Pair<Rule>) -> Option<String> {
    let rec_of = rec_groups_of_module(module)?;
    let (types, names) = descriptor_type_table(module);
    // A target outside the table has no group, and the spec reports that as
    // the rec-group violation it is — `(type (descriptor 1) (struct))` alone
    // asserts "descriptor type is outside rec group", not "unknown type".
    let group = |i: usize| rec_of.get(i).copied();
    for (i, t) in types.iter().enumerate() {
        let own = group(i);
        if let Some(d) = t.descriptor {
            if group(d) != own {
                return Some("descriptor type is outside rec group".to_string());
            }
        }
        if let Some(s) = t.describes {
            if group(s) != own {
                return Some("described type is outside rec group".to_string());
            }
        }
        if t.descriptor.is_some() && !t.is_struct {
            return Some("descriptor type must be a struct".to_string());
        }
        if t.describes.is_some() && !t.is_struct {
            return Some("described type must be a struct".to_string());
        }
        // A descriptor may only describe a type declared BEFORE it — including
        // not itself, which is why this is `>=`.
        if let Some(s) = t.describes {
            if s >= i {
                return Some("forward use of described type".to_string());
            }
        }
        if let Some(d) = t.descriptor {
            if types.get(d).and_then(|x| x.describes) != Some(i) {
                return Some("type is not described by its descriptor".to_string());
            }
        }
        if let Some(s) = t.describes {
            if types.get(s).and_then(|x| x.descriptor) != Some(i) {
                return Some("described type is not described by descriptor".to_string());
            }
        }
    }

    // ── Subtyping: descriptor PRESENCE, then descriptor AGREEMENT ────────────
    //
    // These need no structural subtyping — only the DECLARED `sub` chain — so
    // they are here rather than waiting on the stack-typing pass that the 2297
    // "type mismatch" assertions need. The failure mode is the safe one: a
    // subtyping violation with no descriptor component leaves both sides
    // `None` and this returns None, so a non-descriptor error can never be
    // reported with a descriptor message.
    //
    // ⚠ The two passes are SEPARATE and presence goes first. The
    // "supertype of a descriptor must describe the supertype of the
    // descriptor's described type" fixture violates presence AND agreement at
    // once, and the spec reports the presence failure — running agreement
    // first would answer "descriptor type N does not match" where the suite
    // asserts "sub type 3 does not match super type 1".
    let declared_subtype = |sub: usize, sup: usize| -> bool {
        let mut seen: std::collections::HashSet<usize> = std::collections::HashSet::new();
        let mut stack = vec![sub];
        while let Some(x) = stack.pop() {
            if x == sup {
                return true;
            }
            if !seen.insert(x) {
                continue;
            }
            if let Some(t) = types.get(x) {
                stack.extend(t.supers.iter().copied());
            }
        }
        false
    };
    for (i, t) in types.iter().enumerate() {
        for &sup in &t.supers {
            let Some(st) = types.get(sup) else { continue };
            // A descriptor is inherited downwards but not upwards: if the
            // SUPERTYPE has one the subtype must too, while a subtype may
            // introduce a descriptor its supertype does not have. Only this
            // direction is an error.
            if st.descriptor.is_some() && t.descriptor.is_none() {
                return Some(format!("sub type {i} does not match super type {sup}"));
            }
            // Being a descriptor, unlike having one, must match in BOTH
            // directions — a descriptor may only be a subtype of a descriptor.
            if t.describes.is_some() != st.describes.is_some() {
                return Some(format!("sub type {i} does not match super type {sup}"));
            }
        }
    }
    for t in types.iter() {
        for &sup in &t.supers {
            let Some(st) = types.get(sup) else { continue };
            // The subtype's descriptor must itself be a subtype of the
            // supertype's descriptor. Named by the DESCRIPTOR's index, not the
            // type that carries it.
            if let (Some(d), Some(sd)) = (t.descriptor, st.descriptor) {
                if !declared_subtype(d, sd) {
                    return Some(format!("descriptor type {d} does not match"));
                }
            }
            // Mirror: the supertype's described type must be a supertype of
            // the subtype's described type.
            if let (Some(b), Some(sb)) = (t.describes, st.describes) {
                if !declared_subtype(b, sb) {
                    return Some(format!("described type {b} does not match"));
                }
            }
        }
    }

    // ── Structural agreement of a descriptor type with its supertype ────────
    //
    // ── The GENERAL subtyping rules, before any field comparison ────────────
    //
    // ⛔ ORDER IS THE SPEC'S: a declaration that may not be extended AT ALL is
    // reported as such, not as a field disagreement. Interleaving these with
    // the field walk answers a "sub type" fixture with the wrong reason.
    for (i, t) in types.iter().enumerate() {
        for &sup in &t.supers {
            let Some(st) = types.get(sup) else { continue };
            // ⛔ FINALITY IS THE DEFAULT. `(type $t (struct))` abbreviates
            // `(sub final (struct))`, so a plain declaration is CLOSED and
            // `(type $s (sub $t (struct)))` is invalid. Openness is opt-in.
            if !st.has_sub || st.is_final {
                return Some(format!("sub type {i} does not match super type {sup}"));
            }
            // A subtype must have the same FORM as its supertype — a struct
            // may not extend a func.
            if let (Some(a), Some(b)) = (t.kind, st.kind) {
                if a != b {
                    return Some(format!("sub type {i} does not match super type {sup}"));
                }
            }
            // ⛔ FUNCTION SUBTYPING IS CONTRAVARIANT IN PARAMS, COVARIANT IN
            // RESULTS, and the ARITIES must match exactly — `(func)` extended
            // by `(func (param i32))` is invalid on arity alone. Getting the
            // direction backwards would accept unsound overrides, so the two
            // loops deliberately pass their operands in opposite orders.
            if t.kind == Some("func") && st.kind == Some("func") {
                if let (Some((tp, tr)), Some((sp, sr))) = (&t.func_sig, &st.func_sig) {
                    if tp.len() != sp.len() || tr.len() != sr.len() {
                        return Some(format!("sub type {i} does not match super type {sup}"));
                    }
                    // params: the SUPER's must be usable where the sub's is.
                    for (a, b) in tp.iter().zip(sp.iter()) {
                        if field_incompatible(b, a, &types, &names).is_some() {
                            return Some(format!("sub type {i} does not match super type {sup}"));
                        }
                    }
                    // results: the SUB's must be usable where the super's is.
                    for (a, b) in tr.iter().zip(sr.iter()) {
                        if field_incompatible(a, b, &types, &names).is_some() {
                            return Some(format!("sub type {i} does not match super type {sup}"));
                        }
                    }
                }
            }
            // An ARRAY's element type is its single field. Immutable is
            // covariant, mutable is INVARIANT — same as a struct field.
            if t.kind == Some("array") && st.kind == Some("array") {
                if let (Some(a), Some(b)) = (t.field_types.first(), st.field_types.first()) {
                    if let Some(m) = field_incompatible(a, b, &types, &names) {
                        let _ = m;
                        return Some(format!("sub type {i} does not match super type {sup}"));
                    }
                }
            }
        }
    }

    // A subtype's struct fields must extend its supertype's: it may add
    // fields, but the ones it shares have to agree.
    //
    // ⛔ SCOPED TO DESCRIPTOR TYPES ON PURPOSE. Field compatibility is a
    // GENERAL GC subtyping rule and this helper runs over every module — an
    // imprecise version would fire on the many valid struct hierarchies in the
    // gc suites, which is an overfire in files that have nothing to do with
    // descriptors. Gating on one side of the relationship carrying a
    // descriptor/describes clause keeps it to the rules this file owns; the
    // general case belongs to the typed pass.
    for (i, t) in types.iter().enumerate() {
        for &sup in &t.supers {
            let Some(st) = types.get(sup) else { continue };
            if !t.is_struct || !st.is_struct {
                continue;
            }
            if t.field_types.len() < st.field_types.len() {
                return Some(format!("sub type {i} does not match super type {sup}"));
            }
            for (a, b) in t.field_types.iter().zip(st.field_types.iter()) {
                // ⛔ `!vt_subtype` IS NOT PROOF OF MISMATCH — it also means
                // "cannot decide", and reporting on it is a confident wrong
                // answer about a VALID module. `field_incompatible` reports
                // only what it can establish, mutability included.
                if field_incompatible(a, b, &types, &names).is_some() {
                    return Some(format!("sub type {i} does not match super type {sup}"));
                }
            }
        }
    }

    // ── Allocation and descriptor-read instructions ──────────────────────────
    //
    // Whether a type HAS a descriptor decides which allocation instruction is
    // legal for it, and the two errors are not symmetric spellings of one
    // rule — the spec words them differently, so they are checked separately.
    if let Some(m) = descriptor_instr_reason(module, &types, &names) {
        return Some(m);
    }
    descriptor_operand_mismatch(module, &types, &names)
}

/// The instruction-level Custom Descriptors rules.
///
/// Only the descriptor instructions are examined. A wrong type index on any
/// OTHER instruction is left alone: "unknown type" is a general validation
/// diagnostic and claiming it here would discharge assertions this helper was
/// never written for.
fn descriptor_instr_reason(
    pair: &Pair<Rule>,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<String> {
    if matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
        let mut name: Option<String> = None;
        let mut args: Vec<String> = Vec::new();
        for c in pair.clone().into_inner() {
            match c.as_rule() {
                Rule::instr_name if name.is_none() => name = Some(c.as_str().trim().to_string()),
                Rule::instr_arg => args.push(c.as_str().trim().to_string()),
                _ => {}
            }
        }
        let first_arg = args.first().cloned();
        if let Some(n) = name.as_deref() {
            let base = n.split_once("@@").map(|(b, _)| b).unwrap_or(n);
            let descriptor_instr = matches!(
                base,
                "struct.new"
                    | "struct.new_default"
                    | "struct.new_desc"
                    | "struct.new_default_desc"
                    | "ref.get_desc"
            );
            if descriptor_instr {
                // The operand may be a numeric index or a `$name`. An operand
                // that is neither — a folded sub-expression sitting where the
                // immediate would be — means the immediate is elsewhere, so
                // this instruction is left unjudged.
                let target = first_arg.as_deref().and_then(|a| {
                    a.parse::<usize>()
                        .ok()
                        .or_else(|| names.get(a).copied())
                });
                if let Some(t) = target {
                    let Some(info) = types.get(t) else {
                        // Only the descriptor-read form asserts this; the
                        // allocation forms' out-of-range cases are covered by
                        // the general type-index rules.
                        return (base == "ref.get_desc").then(|| "unknown type".to_string());
                    };
                    let has_desc = info.descriptor.is_some();
                    match base {
                        // ⛔ ASYMMETRIC ON PURPOSE. A type that HAS a
                        // descriptor must be allocated with the `_desc` form,
                        // and a type WITHOUT one must not be — but the spec
                        // words the two failures differently, so they cannot
                        // share a message.
                        "struct.new" | "struct.new_default" if has_desc => {
                            return Some(
                                "type with descriptor requires descriptor allocation".to_string(),
                            );
                        }
                        "struct.new_desc" | "struct.new_default_desc" if !has_desc => {
                            return Some(
                                "type without descriptor requires non-descriptor allocation"
                                    .to_string(),
                            );
                        }
                        // `ref.get_desc` reads a type's descriptor, so the
                        // type must have one. A DESCRIPTOR is not itself
                        // described, so `ref.get_desc` on the descriptor half
                        // of a pair fails here too.
                        "ref.get_desc" if !has_desc => {
                            return Some("type without descriptor".to_string());
                        }
                        _ => {}
                    }
                }
            }
            // The BRANCHING and cast forms take the target as a REFTYPE
            // immediate rather than a bare type index, and they report a
            // missing descriptor differently again — naming the heap type,
            // which for an abstract target is a spelling and not an index.
            //
            //   br_on_cast_desc_eq $l rt_1 rt_2   → rt_2 is the target (arg 2)
            //   ref.cast_desc_eq rt               → arg 0
            let target_arg = match base {
                "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail" => args.get(2),
                "ref.cast_desc_eq" => args.first(),
                _ => None,
            };
            if let Some(rt) = target_arg {
                if let Some(reason) = cast_target_descriptor_reason(rt, types, names) {
                    return Some(reason);
                }
            }
        }
    }
    for child in pair.clone().into_inner() {
        if let Some(r) = descriptor_instr_reason(&child, types, names) {
            return Some(r);
        }
    }
    None
}

// ── Descriptor OPERAND typing ────────────────────────────────────────────────
//
// ⛔ THIS IS NOT A WASM VALIDATOR AND MUST NEVER GROW INTO A HALF-BUILT ONE.
// It types the operands of the DESCRIPTOR instructions and nothing else, and
// it ABANDONS a function the moment it meets anything it cannot type — an
// unmodelled instruction, an unresolvable index, a plain (non-folded)
// instruction whose operands come off a stack it does not track.
//
// That bail is the entire safety property. `assert_invalid` compares the
// expected diagnostic, so a pass that GUESSES reports "type mismatch" on
// modules it never actually typed, and the assertions it turns green mean
// nothing. Unknown must propagate, never default.
//
// The label-typing fixtures ("No types at label", "Too many types at label",
// "Label types do not match fallthrough types") are deliberately NOT covered:
// they need block signatures and a control stack, which is the general pass
// this file does not contain.

#[derive(Clone, PartialEq, Debug)]
enum Heap {
    Abs(&'static str),
    Concrete(usize),
}

#[derive(Clone, PartialEq, Debug)]
struct RefT {
    nullable: bool,
    exact: bool,
    heap: Heap,
}

#[derive(Clone, PartialEq, Debug)]
enum Vt {
    /// ⛔ CARRIES ITS SPELLING. Collapsing the numeric types into one variant
    /// made `i64` a subtype of `i32`, which silently defeated the struct-field
    /// comparison — the fixtures differ by exactly that.
    Num(&'static str),
    Ref(RefT),
    /// `unreachable` — a subtype of everything.
    Bottom,
}

/// The abstract heap types, by their spec spelling. `None` for anything else,
/// which is what makes an unrecognised spelling abandon the function rather
/// than be treated as some default.
fn abs_heap(name: &str) -> Option<&'static str> {
    Some(match name {
        // ⛔ THE `-ref` ABBREVIATIONS ARE THE SAME HEAP TYPE. §2.3.4:
        // `funcref` ≡ `(ref null func)`, `externref` ≡ `(ref null extern)`.
        // `bare_heap_type` lists them FIRST in the grammar, so they reach here
        // — and fell through to `None`, which `ref.null` reads as "concrete
        // type I cannot resolve" and emits as `none`. `ref.null funcref` was
        // therefore a nullref, in a different hierarchy from the funcref it
        // spells. Both readers of this function want the same answer, so the
        // abbreviation resolves here rather than at either call site.
        "funcref" => "func",
        "externref" => "extern",
        // ⛔ `abs_subtype` ALREADY KNOWS `noexn <: exn`; ONLY THIS SIDE DID NOT
        // KNOW THE SPELLING. So the hierarchy was right and unreachable:
        // `parse_vt("exnref")` returned `None`, every exception-typing rule
        // bailed, and the bail is invisible by construction. Same two-helper
        // gap as `funcref`/`externref` above — fixed here, in the shared
        // helper, rather than at either call site.
        "exnref" => "exn",
        "nullexnref" => "noexn",
        "exn" => "exn",
        "noexn" => "noexn",
        "any" => "any",
        "eq" => "eq",
        "i31" => "i31",
        "struct" => "struct",
        "array" => "array",
        "none" => "none",
        "func" => "func",
        "nofunc" => "nofunc",
        "extern" => "extern",
        "noextern" => "noextern",
        _ => return None,
    })
}

/// Parse a val-type / ref-type SPELLING. Handles `i32`, the §2.3.4
/// abbreviations, `(ref $t)`, `(ref null $t)` and the proposal's
/// `(ref (exact $t))`.
fn parse_vt(text: &str, names: &HashMap<String, usize>) -> Option<Vt> {
    let t = text.trim();
    if let Some(n) = match t {
        "i32" => Some("i32"),
        "i64" => Some("i64"),
        "f32" => Some("f32"),
        "f64" => Some("f64"),
        "v128" => Some("v128"),
        _ => None,
    } {
        return Some(Vt::Num(n));
    }
    if let Some((heap, _)) = vybe_runtime::opcode::heaptype::HeapType::from_spec_reftype_name(t) {
        return Some(Vt::Ref(RefT {
            nullable: true,
            exact: false,
            heap: Heap::Abs(abs_heap(heap)?),
        }));
    }
    // ⛔ THE RUNTIME'S ABBREVIATION TABLE STOPS BEFORE THE EXCEPTION HIERARCHY:
    // it lists `anyref` … `nullexternref` and no `exnref`. That made this
    // function answer `None` for a type the rest of the pass models perfectly —
    // `abs_subtype` already knows `noexn <: exn` — so every exception-typing
    // rule bailed, silently, on the spelling alone.
    //
    // Restricted to names ENDING IN `ref` on purpose. `abs_heap` also maps bare
    // HEAP names (`any`, `struct`, …), and those are not val types: falling
    // back on every name would make `(param any)` parse, which is a wrong
    // answer where `None` was merely an absent one.
    if t.ends_with("ref") {
        if let Some(h) = abs_heap(t) {
            return Some(Vt::Ref(RefT {
                nullable: true,
                exact: false,
                heap: Heap::Abs(h),
            }));
        }
    }
    if !t.starts_with('(') {
        return None;
    }
    let inner = t.trim_start_matches('(').trim_end_matches(')').trim();
    let rest = inner.strip_prefix("ref")?.trim();
    let (rest, nullable) = match rest.strip_prefix("null") {
        Some(r) => (r.trim(), true),
        None => (rest, false),
    };
    // `(exact $t)` — the proposal's exact heap types.
    let (spelling, exact) = match rest.strip_prefix("(exact") {
        Some(r) => (r.trim().trim_end_matches(')').trim(), true),
        None => (rest, false),
    };
    let heap = match abs_heap(spelling) {
        Some(a) => Heap::Abs(a),
        None => Heap::Concrete(
            spelling
                .parse::<usize>()
                .ok()
                .or_else(|| names.get(spelling).copied())?,
        ),
    };
    Some(Vt::Ref(RefT {
        nullable,
        exact,
        heap,
    }))
}

/// The param and result type SPELLINGS under a pair, in order.
fn collect_params_results(p: &Pair<Rule>, ps: &mut Vec<String>, rs: &mut Vec<String>) {
    for c in p.clone().into_inner() {
        match c.as_rule() {
            Rule::param => {
                for d in c.into_inner() {
                    if d.as_rule() != Rule::id {
                        ps.push(d.as_str().trim().to_string());
                    }
                }
            }
            Rule::result => {
                for d in c.into_inner() {
                    rs.push(d.as_str().trim().to_string());
                }
            }
            _ => collect_params_results(&c, ps, rs),
        }
    }
}

/// A field/element spelling with any `(mut …)` wrapper removed.
fn strip_mut(s: &str) -> String {
    let t = s.trim();
    t.strip_prefix('(')
        .and_then(|x| x.trim().strip_prefix("mut"))
        .map(|x| x.trim().trim_end_matches(')').trim().to_string())
        .unwrap_or_else(|| t.to_string())
}

/// Are two field spellings DEMONSTRABLY incompatible as sub/super?
///
/// ⛔ MUTABILITY IS INVARIANT AND ITS PRESENCE MUST MATCH. `(mut T)` may only
/// be extended by `(mut T)` with the SAME `T` — narrowing a mutable field is
/// unsound because it can still be written through the supertype. An immutable
/// field is covariant. Dropping or adding `mut` is a mismatch either way, and
/// the suite asserts exactly that pair.
fn field_incompatible(
    a: &str,
    b: &str,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<&'static str> {
    let unmut = |s: &str| -> Option<String> {
        let t = s.trim();
        let inner = t.strip_prefix('(')?.trim().strip_prefix("mut")?;
        Some(inner.trim().trim_end_matches(')').trim().to_string())
    };
    match (unmut(a), unmut(b)) {
        // Both mutable: INVARIANT, so the spellings must denote the same type.
        (Some(x), Some(y)) => {
            if x == y {
                return None;
            }
            match (parse_vt(&x, names), parse_vt(&y, names)) {
                (Some(px), Some(py)) => {
                    if vt_subtype(&px, &py, types) && vt_subtype(&py, &px, types) {
                        None
                    } else if provably_not_subtype(&px, &py, types)
                        || provably_not_subtype(&py, &px, types)
                    {
                        Some("mutable field is invariant")
                    } else {
                        None
                    }
                }
                _ => None,
            }
        }
        // One mutable, one not — a mismatch whichever way round.
        (Some(_), None) | (None, Some(_)) => Some("mutability differs"),
        (None, None) => {
            if a.trim() == b.trim() {
                return None;
            }
            match (parse_vt(a, names), parse_vt(b, names)) {
                (Some(x), Some(y)) if vt_subtype(&x, &y, types) => None,
                (Some(x), Some(y)) if provably_not_subtype(&x, &y, types) => {
                    Some("field is not a subtype")
                }
                _ => None,
            }
        }
    }
}

/// Is `a` DEMONSTRABLY not a subtype of `b`?
///
/// ⛔ THE COMPLEMENT OF `vt_subtype` IS NOT THIS. `vt_subtype` answers "can I
/// prove yes"; its `false` covers both "provably no" and "cannot tell", and a
/// rule that reports on `false` alone announces a mismatch it never
/// established.
///
/// The undecidable case here is real and the suite pins it: WASM types are
/// ISO-RECURSIVE, so two concrete types in DIFFERENT rec groups whose groups
/// are structurally identical ARE the same type. `gc/type-subtyping.wast`
/// builds exactly that — `(rec (type $f1 …) (type $s1 …))` and
/// `(rec (type $f2 …) (type $s2 …))` with identical shapes, making
/// `(ref $f1)` and `(ref $f2)` interchangeable. Deciding it needs rec-group
/// CANONICALISATION, which this file does not do; until it does, two distinct
/// concrete indices must answer "don't know", never "different".
fn provably_not_subtype(a: &Vt, b: &Vt, types: &[DescType]) -> bool {
    match (a, b) {
        (Vt::Bottom, _) | (_, Vt::Bottom) => false,
        // Different scalar spellings are decidable outright — and this is the
        // case the fixtures actually assert.
        (Vt::Num(x), Vt::Num(y)) => x != y,
        (Vt::Num(_), Vt::Ref(_)) | (Vt::Ref(_), Vt::Num(_)) => true,
        (Vt::Ref(x), Vt::Ref(y)) => match (&x.heap, &y.heap) {
            // Decidable ONCE both canonical ids are known — that is exactly
            // what canonicalisation buys. Without them this stays "don't
            // know", because a `false` from the sub-chain walk alone proves
            // nothing about twin rec groups.
            (Heap::Concrete(i), Heap::Concrete(j)) => {
                match (
                    types.get(*i).and_then(|t| t.canon),
                    types.get(*j).and_then(|t| t.canon),
                ) {
                    (Some(_), Some(_)) => !heap_subtype(&x.heap, &y.heap, types),
                    _ => false,
                }
            }
            // A concrete type's abstract head depends on its form, which this
            // table does not record — undecidable here.
            (Heap::Concrete(_), Heap::Abs(_)) | (Heap::Abs(_), Heap::Concrete(_)) => false,
            // FULLY DECIDABLE now that `abs_subtype` is complete: the abstract
            // lattice is finite and closed.
            (Heap::Abs(p), Heap::Abs(q)) => !abs_subtype(p, q) || (x.nullable && !y.nullable),
        },
    }
}

/// Is `a` a subtype of `b`?
///
/// ⚠ EXACTNESS IS NOT A FLAG THAT WIDENS. `(ref (exact $d))` is NOT a subtype
/// of `(ref (exact $b))` even when `$d` is a subtype of `$b` — an exact target
/// admits that type and nothing under it. The suite pins this directly ("An
/// exact reference to a subtype of the descriptor does not cut it").
fn vt_subtype(a: &Vt, b: &Vt, types: &[DescType]) -> bool {
    match (a, b) {
        (Vt::Bottom, _) => true,
        (Vt::Num(x), Vt::Num(y)) => x == y,
        (Vt::Ref(x), Vt::Ref(y)) => {
            if x.nullable && !y.nullable {
                return false;
            }
            if y.exact {
                return x.exact && x.heap == y.heap;
            }
            heap_subtype(&x.heap, &y.heap, types)
        }
        _ => false,
    }
}

/// The COMPLETE abstract heap-type lattice.
///
/// ⛔ THE PARTIAL VERSION WAS NOT SAFE TO DECIDE ON. It knew `struct/array/i31
/// <: eq` and the `any` roof but not `nofunc <: func` or `noextern <: extern`,
/// so treating its `false` as "provably not a subtype" would have rejected
/// valid modules. Abstract subtyping is finite and fully decidable; write all
/// of it, then it can be relied on.
///
/// Three disjoint hierarchies — internal (`any`), function (`func`) and
/// external (`extern`) — each with its own bottom. Nothing crosses between.
fn abs_subtype(a: &str, b: &str) -> bool {
    if a == b {
        return true;
    }
    let internal = ["any", "eq", "i31", "struct", "array", "none"];
    let is_internal = |t: &str| internal.contains(&t);
    match (a, b) {
        // Bottoms.
        ("none", t) if is_internal(t) => true,
        ("nofunc", "func") => true,
        ("noextern", "extern") => true,
        ("noexn", "exn") => true,
        // Internal roof and the `eq` layer.
        (x, "any") if is_internal(x) => true,
        ("i31" | "struct" | "array", "eq") => true,
        _ => false,
    }
}

fn heap_subtype(a: &Heap, b: &Heap, types: &[DescType]) -> bool {
    match (a, b) {
        (Heap::Abs(x), Heap::Abs(y)) => abs_subtype(x, y),
        // A concrete type sits under the abstract head of its OWN hierarchy,
        // and which head that is depends on its FORM. The table records the
        // form now, so this is exact rather than the old
        // "concrete is below struct and eq" approximation — which both missed
        // `$t <: any` and wrongly put a concrete FUNC type under `struct`.
        (Heap::Concrete(i), Heap::Abs(y)) => match types.get(*i).and_then(|t| t.kind) {
            Some(k) => abs_subtype(k, y),
            // Form unknown: keep the permissive answer rather than invent one.
            None => matches!(*y, "struct" | "eq" | "any"),
        },
        // `none` bottoms the INTERNAL hierarchy only — `nofunc` bottoms
        // functions. A concrete func type is not above `none`.
        (Heap::Abs("none"), Heap::Concrete(j)) => {
            !matches!(types.get(*j).and_then(|t| t.kind), Some("func"))
        }
        (Heap::Abs("nofunc"), Heap::Concrete(j)) => {
            matches!(types.get(*j).and_then(|t| t.kind), Some("func"))
        }
        (Heap::Concrete(i), Heap::Concrete(j)) => {
            // ⛔ SAME TYPE, DIFFERENT DECLARATION. Under iso-recursive
            // equivalence two structurally identical rec groups ARE one type,
            // so the declared-`sub`-chain walk below is not the whole answer —
            // it says "no" for `(ref $f1)` vs `(ref $f2)` in twin groups,
            // which is a valid module.
            match (
                types.get(*i).and_then(|t| t.canon),
                types.get(*j).and_then(|t| t.canon),
            ) {
                (Some(a), Some(b)) if a == b => return true,
                _ => {}
            }
            let mut seen: std::collections::HashSet<usize> = std::collections::HashSet::new();
            let mut stack = vec![*i];
            while let Some(x) = stack.pop() {
                if x == *j {
                    return true;
                }
                if !seen.insert(x) {
                    continue;
                }
                if let Some(t) = types.get(x) {
                    stack.extend(t.supers.iter().copied());
                }
            }
            false
        }
        _ => false,
    }
}

/// The type a folded operand produces, or `None` — which abandons the
/// function. Every arm here is an instruction whose result type is decided by
/// its IMMEDIATE, so no operand stack is needed.
fn infer_operand(
    pair: &Pair<Rule>,
    locals: &[Vt],
    local_names: &HashMap<String, usize>,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<Vt> {
    // Unwrap the non-silent `instr` wrapper if it is still on.
    if pair.as_rule() == Rule::instr {
        return infer_operand(&pair.clone().into_inner().next()?, locals, local_names, types, names);
    }
    if !matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
        return None;
    }
    let mut head: Option<String> = None;
    let mut args: Vec<String> = Vec::new();
    for c in pair.clone().into_inner() {
        match c.as_rule() {
            Rule::instr_name if head.is_none() => head = Some(c.as_str().trim().to_string()),
            Rule::instr_arg => args.push(c.as_str().trim().to_string()),
            _ => {}
        }
    }
    let h = head?;
    let base = h.split_once("@@").map(|(b, _)| b).unwrap_or(&h);
    match base {
        "unreachable" => Some(Vt::Bottom),
        "i32.const" => Some(Vt::Num("i32")),
        "i64.const" => Some(Vt::Num("i64")),
        "f32.const" => Some(Vt::Num("f32")),
        "f64.const" => Some(Vt::Num("f64")),
        "local.get" => {
            let a = args.first()?;
            let idx = a
                .parse::<usize>()
                .ok()
                .or_else(|| local_names.get(a.as_str()).copied())?;
            locals.get(idx).cloned()
        }
        "ref.null" => {
            let a = args.first()?;
            let heap = match abs_heap(a) {
                Some(x) => Heap::Abs(x),
                None => Heap::Concrete(
                    a.parse::<usize>()
                        .ok()
                        .or_else(|| names.get(a.as_str()).copied())?,
                ),
            };
            Some(Vt::Ref(RefT {
                nullable: true,
                exact: false,
                heap,
            }))
        }
        // An allocation yields an EXACT, non-null reference to the type it
        // names — that is what makes a descriptor operand typed by
        // `struct.new $b` acceptable where `(ref (exact $b))` is required.
        "struct.new" | "struct.new_default" | "struct.new_desc" | "struct.new_default_desc" => {
            let a = args.first()?;
            let idx = a
                .parse::<usize>()
                .ok()
                .or_else(|| names.get(a.as_str()).copied())?;
            let _ = types.get(idx)?;
            Some(Vt::Ref(RefT {
                nullable: false,
                exact: true,
                heap: Heap::Concrete(idx),
            }))
        }
        "ref.cast" => parse_vt(args.first()?, names),
        _ => None,
    }
}

/// Split an instruction's children into IMMEDIATES (text) and OPERANDS (folded
/// sub-instructions), in source order.
///
/// ⛔ `instr_arg` can itself BE a `folded_instr` — the grammar alternation puts
/// them in the same slot — so an operand and an immediate are not told apart
/// by their parent rule. The inner rule is what decides.
// ⛔ `&Pair<'a, …>`, NOT `&'a Pair<'a, …>`. A pest `Pair<'a>` is Clone and its
// `'a` is the lifetime of the PARSED INPUT, not of the borrow — tying the two
// together forced every caller to keep the borrowed pair alive for as long as
// the returned operands, which a loop over owned pairs cannot do. Written the
// tight way first; it compiled only because my own callers happened to hold
// the pair, and broke the first caller that did not.
fn split_immediates_and_operands<'a>(
    pair: &Pair<'a, Rule>,
) -> (Vec<String>, Vec<Pair<'a, Rule>>) {
    let mut imms = Vec::new();
    let mut ops = Vec::new();
    for c in pair.clone().into_inner() {
        match c.as_rule() {
            Rule::instr_arg => match c.clone().into_inner().next() {
                Some(inner) if inner.as_rule() == Rule::folded_instr => ops.push(inner),
                _ => imms.push(c.as_str().trim().to_string()),
            },
            // ⛔ `instr = { folded_instr | plain_instr }` is NOT silent, so an
            // operand arrives wrapped and never as a bare `folded_instr`.
            // Matching only the bare forms found zero operands and the whole
            // pass sat inert while still compiling.
            Rule::instr => {
                if let Some(inner) = c.into_inner().next() {
                    ops.push(inner);
                }
            }
            Rule::folded_instr | Rule::plain_instr => ops.push(c),
            _ => {}
        }
    }
    (imms, ops)
}

/// The declared local types of a function, params first — `None` if any
/// declaration cannot be parsed, which abandons the function.
fn func_locals(
    f: &Pair<Rule>,
    names: &HashMap<String, usize>,
) -> Option<(Vec<Vt>, HashMap<String, usize>, Vec<bool>)> {
    let mut out = Vec::new();
    let mut by_name = HashMap::new();
    // §3.4.1: a local starts INITIALIZED only if its type is defaultable.
    // Parameters always are — they arrive with a value — so the two decl kinds
    // answer differently and `collect_decls` keeps them distinguishable.
    let mut inited: Vec<bool> = Vec::new();
    // ⛔ `param` is NOT a direct child of `func_field` — it sits inside
    // `typeuse` (`func_field = { … ~ typeuse ~ … ~ local* ~ instr* }`).
    // Iterating one level found no params at all, so every function typed as
    // having zero locals and the pass sat inert while compiling and passing.
    // Collected in document order, which is the index order params and locals
    // share.
    let mut decls: Vec<Pair<Rule>> = Vec::new();
    fn collect_decls<'a>(p: Pair<'a, Rule>, out: &mut Vec<Pair<'a, Rule>>) {
        if matches!(p.as_rule(), Rule::param | Rule::local) {
            out.push(p);
            return;
        }
        // Do not descend into a nested function type: its params are not this
        // function's locals.
        if p.as_rule() == Rule::func_type {
            return;
        }
        for c in p.into_inner() {
            collect_decls(c, out);
        }
    }
    for c in f.clone().into_inner() {
        collect_decls(c, &mut decls);
    }
    for c in decls {
        let mut id: Option<String> = None;
        let mut decls: Vec<String> = Vec::new();
        for d in c.clone().into_inner() {
            match d.as_rule() {
                Rule::id => id = Some(d.as_str().trim().to_string()),
                _ => decls.push(d.as_str().trim().to_string()),
            }
        }
        if let Some(n) = id {
            by_name.insert(n, out.len());
        }
        let is_param = c.as_rule() == Rule::param;
        for d in decls {
            let t = parse_vt(&d, names)?;
            // Defaultable ⇔ not a non-nullable reference.
            let defaultable = !matches!(&t, Vt::Ref(r) if !r.nullable);
            inited.push(is_param || defaultable);
            out.push(t);
        }
    }
    Some((out, by_name, inited))
}

/// Operand typing for the descriptor instructions.
///
/// Returns `None` when the function could not be fully typed — the bail — and
/// `Some(msg)` only for a mismatch it actually proved.
fn descriptor_operand_mismatch(
    module: &Pair<Rule>,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<String> {
    for c in module.clone().into_inner() {
        if c.as_rule() != Rule::module_field {
            continue;
        }
        let Some(f) = c.into_inner().next() else { continue };
        if f.as_rule() != Rule::func_field {
            continue;
        }
        let Some((locals, local_names, _)) = func_locals(&f, names) else {
            continue;
        };
        if let Some(m) = scan_descriptor_operands(&f, &locals, &local_names, types, names) {
            return Some(m);
        }
    }
    None
}

fn scan_descriptor_operands(
    pair: &Pair<Rule>,
    locals: &[Vt],
    local_names: &HashMap<String, usize>,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<String> {
    if matches!(pair.as_rule(), Rule::folded_instr) {
        let head = pair
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::instr_name)
            .map(|c| c.as_str().trim().to_string());
        if let Some(h) = head {
            let base = h.split_once("@@").map(|(b, _)| b).unwrap_or(&h);
            let (imms, ops) = split_immediates_and_operands(pair);
            // Which immediate names the target, and how many operands the
            // instruction takes, is per-instruction — there is no general rule
            // to lean on here.
            let target_spelling = match base {
                "ref.cast_desc_eq" => imms.first(),
                "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail" => imms.get(2),
                _ => None,
            };
            // ⚠ IMMEDIATE-ONLY, so it holds even when the operand is
            // `unreachable`. `br_on_cast_desc_eq $l rt_1 rt_2` requires
            // rt_2 <: rt_1; `(ref null func)` with `(ref $a)` is two
            // hierarchies and the suite asserts that with a bottom operand
            // precisely so no operand typing can excuse it.
            if matches!(base, "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail") {
                if let (Some(a), Some(b)) = (imms.get(1), imms.get(2)) {
                    if let (Some(from), Some(to)) = (parse_vt(a, names), parse_vt(b, names)) {
                        if !vt_subtype(&to, &from, types) {
                            return Some("type mismatch".to_string());
                        }
                    }
                }
            }
            if let Some(rt) = target_spelling {
                if let Some(Vt::Ref(target)) = parse_vt(rt, names) {
                    if let Heap::Concrete(t) = target.heap {
                        if let Some(desc_idx) = types.get(t).and_then(|x| x.descriptor) {
                            // The descriptor operand is the LAST one, and its
                            // exactness is inherited from the cast: an exact
                            // cast demands an exact descriptor.
                            if let Some(dop) = ops.last() {
                                let Some(got) =
                                    infer_operand(dop, locals, local_names, types, names)
                                else {
                                    return None;
                                };
                                let want = Vt::Ref(RefT {
                                    nullable: true,
                                    exact: target.exact,
                                    heap: Heap::Concrete(desc_idx),
                                });
                                if !vt_subtype(&got, &want, types) {
                                    return Some("type mismatch".to_string());
                                }
                            }
                            // The CAST VALUE has to be in the same type
                            // hierarchy as the target. A descriptor target is
                            // always a struct, so the value must sit under
                            // `any` — `(ref null func)` is a different
                            // hierarchy and the suite calls that out directly
                            // ("Cannot cast across hierarchies").
                            if ops.len() >= 2 {
                                let Some(got) =
                                    infer_operand(&ops[0], locals, local_names, types, names)
                                else {
                                    return None;
                                };
                                let any = Vt::Ref(RefT {
                                    nullable: true,
                                    exact: false,
                                    heap: Heap::Abs("any"),
                                });
                                if !vt_subtype(&got, &any, types) {
                                    return Some("type mismatch".to_string());
                                }
                            }
                        }
                    }
                }
            }
            // Allocation arity: a `_desc` form takes the struct's fields PLUS
            // one descriptor. Counted only in the folded spelling, where the
            // operands are syntactically present.
            if matches!(
                base,
                "struct.new_desc" | "struct.new_default_desc" | "struct.new" | "struct.new_default"
            ) {
                if let Some(a) = imms.first() {
                    if let Some(idx) = a
                        .parse::<usize>()
                        .ok()
                        .or_else(|| names.get(a.as_str()).copied())
                    {
                        if let Some(info) = types.get(idx) {
                            // `_desc` forms take one more operand than their
                            // plain counterparts: the descriptor.
                            let want = match base {
                                "struct.new_desc" => info.fields + 1,
                                "struct.new_default_desc" => 1,
                                "struct.new" => info.fields,
                                _ => 0,
                            };
                            if ops.len() != want {
                                return Some("type mismatch".to_string());
                            }
                        }
                    }
                }
            }
            // `ref.get_desc $t` reads the descriptor OFF a `$t`, so its
            // operand must actually be one.
            if base == "ref.get_desc" {
                if let (Some(a), Some(op)) = (imms.first(), ops.first()) {
                    if let Some(idx) = a
                        .parse::<usize>()
                        .ok()
                        .or_else(|| names.get(a.as_str()).copied())
                    {
                        if types.get(idx).is_some() {
                            let Some(got) =
                                infer_operand(op, locals, local_names, types, names)
                            else {
                                return None;
                            };
                            let want = Vt::Ref(RefT {
                                nullable: true,
                                exact: false,
                                heap: Heap::Concrete(idx),
                            });
                            if !vt_subtype(&got, &want, types) {
                                return Some("type mismatch".to_string());
                            }
                        }
                    }
                }
            }
        }
    }
    for child in pair.clone().into_inner() {
        if let Some(m) = scan_descriptor_operands(&child, locals, local_names, types, names) {
            return Some(m);
        }
    }
    None
}

/// A descriptor cast's TARGET reftype must name a type that has a descriptor.
///
/// The diagnostic spells the target the way the SOURCE does: a concrete target
/// is reported by index ("type 0 does not have a descriptor"), an abstract one
/// by its heap type ("type any …", "type none …") — so `anyref` reports as
/// `any` and `nullref` as `none`, via the §2.3.4 abbreviation table rather
/// than a list kept here.
fn cast_target_descriptor_reason(
    arg: &str,
    types: &[DescType],
    names: &HashMap<String, usize>,
) -> Option<String> {
    // `(ref null $a)` / `(ref 1)` / `anyref`. The heap type is the last token
    // once the reftype wrapper and its `null` marker are removed.
    let inner = arg.trim().trim_start_matches('(').trim_end_matches(')');
    let spelling = inner
        .split_whitespace()
        .filter(|t| *t != "ref" && *t != "null")
        .next_back()?;
    let concrete = spelling
        .parse::<usize>()
        .ok()
        .or_else(|| names.get(spelling).copied());
    match concrete {
        Some(n) => match types.get(n) {
            None => Some("unknown type".to_string()),
            Some(t) if t.descriptor.is_none() => {
                Some(format!("type {n} does not have a descriptor"))
            }
            Some(_) => None,
        },
        None => {
            // An abstract heap type never has a descriptor. Anything that is
            // not a recognised spelling is left alone rather than guessed at.
            let heap = vybe_runtime::opcode::heaptype::HeapType::from_spec_reftype_name(spelling)
                .map(|(h, _)| h)
                .or(
                    match vybe_runtime::opcode::heaptype::HeapType::from_spec_name(spelling) {
                        Some(_) => Some(spelling),
                        None => None,
                    },
                )?;
            Some(format!("type {heap} does not have a descriptor"))
        }
    }
}

/// What a module DECLARES, per entity kind: the set of `$id`s and the total
/// count. A reference resolves against one or the other — `$name` against the
/// set, a bare integer against the count — which is exactly the pair the
/// "unknown X" diagnostics test.
#[derive(Default)]
struct ModuleCensus {
    types: (std::collections::HashSet<String>, usize),
    funcs: (std::collections::HashSet<String>, usize),
    globals: (std::collections::HashSet<String>, usize),
    data_segs: (std::collections::HashSet<String>, usize),
    elem_segs: (std::collections::HashSet<String>, usize),
    // ⛔ COUNTED, NOT NAMED, WAS A BUG. `(export "a" (table $a))` resolved
    // `$a` against `Default::default()` — an EMPTY name set — so every
    // by-name export of a table or memory reported "unknown table" on a
    // perfectly valid module. Nothing in the suite could show it: these rules
    // only ever run inside `assert_invalid`, where the module is asserted
    // invalid already. The false-positive sweep over VALID modules is what
    // surfaced it.
    memories: (std::collections::HashSet<String>, usize),
    tables: (std::collections::HashSet<String>, usize),
    tags: (std::collections::HashSet<String>, usize),
}

fn census_id(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
}

/// Count every declaration in a module, both spellings of an import included.
///
/// ⛔ IMPORTS DECLARE ENTITIES. `(import "m" "n" (func $f …))` and
/// `(func $f (import "m" "n") …)` each add a function; counting only
/// `func_field` makes `call 3` in a module with three imports look out of
/// range and reports "unknown function" for a call that is perfectly fine.
/// Rec-group members splice in as ordinary `module_field`s, so no special case.
fn build_census(module: &Pair<Rule>, c: &mut ModuleCensus) {
    for pair in module.clone().into_inner() {
        let inner = match pair.as_rule() {
            Rule::module_field => match pair.clone().into_inner().next() {
                Some(i) => i,
                None => continue,
            },
            _ => pair.clone(),
        };
        match inner.as_rule() {
            Rule::type_field => {
                if let Some(id) = census_id(&inner) {
                    c.types.0.insert(id);
                }
                c.types.1 += 1;
            }
            Rule::func_field => {
                if let Some(id) = census_id(&inner) {
                    c.funcs.0.insert(id);
                }
                c.funcs.1 += 1;
            }
            Rule::global_field => {
                if let Some(id) = census_id(&inner) {
                    c.globals.0.insert(id);
                }
                c.globals.1 += 1;
            }
            Rule::data_field => {
                if let Some(id) = census_id(&inner) {
                    c.data_segs.0.insert(id);
                }
                c.data_segs.1 += 1;
            }
            Rule::elem_field => {
                if let Some(id) = census_id(&inner) {
                    c.elem_segs.0.insert(id);
                }
                c.elem_segs.1 += 1;
            }
            Rule::memory_field => {
                if let Some(id) = census_id(&inner) {
                    c.memories.0.insert(id);
                }
                c.memories.1 += 1;
            }
            Rule::table_field => {
                if let Some(id) = census_id(&inner) {
                    c.tables.0.insert(id);
                }
                c.tables.1 += 1;
            }
            Rule::tag_field => {
                if let Some(id) = census_id(&inner) {
                    c.tags.0.insert(id);
                }
                c.tags.1 += 1;
            }
            Rule::import_field => {
                if let Some(desc) = inner
                    .clone()
                    .into_inner()
                    .find(|x| x.as_rule() == Rule::import_desc)
                {
                    let id = census_id(&desc);
                    // The descriptor's keyword is its first token.
                    let kind = desc.as_str().trim_start_matches('(').trim_start();
                    let kind = kind.split_whitespace().next().unwrap_or("");
                    // An IMPORTED memory or table declares one just as a
                    // defined field does — `(import "m" "n" (memory 1))` makes
                    // memory 0 exist, so a load in that module is fine.
                    match kind {
                        "memory" => {
                            if let Some(n) = census_id(&desc) {
                                c.memories.0.insert(n);
                            }
                            c.memories.1 += 1;
                        }
                        "table" => {
                            if let Some(n) = census_id(&desc) {
                                c.tables.0.insert(n);
                            }
                            c.tables.1 += 1;
                        }
                        "tag" => {
                            if let Some(n) = census_id(&desc) {
                                c.tags.0.insert(n);
                            }
                            c.tags.1 += 1;
                        }
                        _ => {}
                    }
                    let slot = match kind {
                        "func" => Some(&mut c.funcs),
                        "global" => Some(&mut c.globals),
                        _ => None,
                    };
                    if let Some(slot) = slot {
                        if let Some(id) = id {
                            slot.0.insert(id);
                        }
                        slot.1 += 1;
                    }
                }
            }
            _ => {}
        }
    }
    c.types.1 += implicit_type_upper_bound(module);
}

/// How many types a module's inline signatures can ADD to the type index
/// space, as an upper bound.
///
/// ⛔ THE TYPE SECTION IS NOT JUST THE `(type …)` FIELDS. §6.6.4: a `typeuse`
/// written inline — `(func $f (result f64) …)` — DEFINES a type when the
/// module has no matching one, and it lands in the SAME index space. A module
/// with one explicit `(type $t …)` and two inline signatures therefore has
/// three types, and `(type 1)` in it is perfectly legal. Counting only the
/// explicit fields reported "unknown type" on the suite's own func.wast — a
/// VALID module, which is why no `assert_invalid` fixture could ever show it.
///
/// An UPPER bound is the right shape: the spec DEDUPES, so the true count is
/// at most this. Over-counting can only make us miss a genuine out-of-range
/// index, leaving that assertion honestly red; under-counting invents an error
/// on a valid module, which is the failure this exists to prevent.
fn implicit_type_upper_bound(pair: &Pair<Rule>) -> usize {
    let mut n = 0;
    if pair.as_rule() == Rule::typeuse {
        let has_index = pair.clone().into_inner().any(|c| c.as_rule() == Rule::index);
        let has_sig = pair
            .clone()
            .into_inner()
            .any(|c| matches!(c.as_rule(), Rule::param | Rule::result));
        if !has_index && has_sig {
            n += 1;
        }
    }
    for c in pair.clone().into_inner() {
        n += implicit_type_upper_bound(&c);
    }
    n
}

/// Does `idx` (an `index` pair: `$name` or an integer) resolve?
fn index_resolves(idx: &Pair<Rule>, decl: &(std::collections::HashSet<String>, usize)) -> bool {
    let t = idx.as_str().trim();
    match t.strip_prefix('$') {
        Some(name) => decl.0.contains(name),
        None => match t.parse::<usize>() {
            Ok(n) => n < decl.1,
            // Not a resolvable spelling at all — leave it to another check
            // rather than inventing a diagnostic for it.
            Err(_) => true,
        },
    }
}

/// The first instruction inside a constant-expression context that is not a
/// constant form, if any.
///
/// ⛔ The set is the SPEC's, not "things that happen to fold": `i32.ctz` is
/// perfectly computable at compile time and is still not a constant expression.
/// `i32.add`/`sub`/`mul` (and the i64 forms) ARE, but only because the
/// extended-const proposal is merged into 3.0.
fn first_non_const_instr(pair: &Pair<Rule>) -> Option<String> {
    const CONST_OPS: &[&str] = &[
        "i32.const", "i64.const", "f32.const", "f64.const", "v128.const",
        "ref.null", "ref.func", "ref.i31",
        "global.get",
        "struct.new", "struct.new_default", "struct.new_desc", "struct.new_default_desc",
        "array.new", "array.new_default", "array.new_fixed",
        // extended-const (merged in WASM 3.0)
        "i32.add", "i32.sub", "i32.mul", "i64.add", "i64.sub", "i64.mul",
        "any.convert_extern", "extern.convert_any",
    ];
    fn walk(p: Pair<Rule>, out: &mut Option<String>) {
        if out.is_some() {
            return;
        }
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            if let Some(n) = instr_head_name(&p) {
                if !CONST_OPS.contains(&n.as_str()) {
                    *out = Some(n);
                    return;
                }
            }
        }
        for ch in p.into_inner() {
            walk(ch, out);
        }
    }
    let mut out = None;
    walk(pair.clone(), &mut out);
    out
}

/// Is this `elem_mode` actually the segment's REFERENCE TYPE, mis-matched as an
/// offset by the grammar's ordering? An offset computes an address; a `ref.*`
/// form cannot.
fn elem_mode_is_reference_type(mode: &Pair<Rule>) -> bool {
    fn head(p: Pair<Rule>) -> Option<String> {
        if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            return instr_head_name(&p);
        }
        p.into_inner().find_map(head)
    }
    // ⛔ `(ref 1)` folds to an instruction named plain `ref` — no dot — so
    // matching only `ref.` missed the very shape this exists for.
    head(mode.clone()).is_some_and(|n| n == "ref" || n.starts_with("ref."))
}

/// The mnemonic at the head of an instruction pair, suffixes stripped.
fn instr_head_name(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::instr_name)
        .map(|c| {
            let n = c.as_str();
            n.split_once("@@").map(|(b, _)| b).unwrap_or(n).to_string()
        })
}

/// Does this instruction access linear memory? Every load/store declares a
/// natural alignment — that IS the property of touching memory — and the
/// `memory.*` family does so by name. Asked of the OPCODE TABLE rather than a
/// list kept here.
fn instr_touches_memory(name: &str) -> bool {
    if name.starts_with("memory.") {
        return true;
    }
    vybe_runtime::opcode::Op::from_wasm_name(name)
        .and_then(|op| op.natural_align_bytes())
        .is_some()
}

/// Does this instruction access a table? `call_indirect` needs one even though
/// its name does not say so.
fn instr_touches_table(name: &str) -> bool {
    name.starts_with("table.") || matches!(name, "call_indirect" | "return_call_indirect")
}

/// A WAT `uN` as a u128 — decimal or `0x` hex, `_` separators. Wide on purpose:
/// the limits fixtures write `0x1_0000_0000`, which overflows the u32 the value
/// is stored in, and the check needs to SEE the overflow rather than wrap.
/// Returns None for anything that is not a well-formed unsigned literal, which
/// is the malformed half's business.
fn parse_wat_u128(text: &str) -> Option<u128> {
    let t = text.trim().replace('_', "");
    match t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
        Some(hex) => u128::from_str_radix(hex, 16).ok(),
        None => t.parse::<u128>().ok(),
    }
}

/// The number of lanes an instruction's lane immediate indexes, read off the
/// MNEMONIC — which is where WASM states the shape. `i16x8.extract_lane` has 8;
/// `v128.store64_lane` writes 64 bits of a 128-bit vector, so 2.
fn mnemonic_lane_count(name: &str) -> Option<u32> {
    if !name.contains("_lane") {
        return None;
    }
    if let Some(shape) = name.split('.').next() {
        match shape {
            "i8x16" => return Some(16),
            "i16x8" => return Some(8),
            "i32x4" | "f32x4" => return Some(4),
            "i64x2" | "f64x2" => return Some(2),
            _ => {}
        }
    }
    // `v128.loadN_lane` / `v128.storeN_lane` — N bits per lane.
    for (tag, lanes) in [("8_lane", 16u32), ("16_lane", 8), ("32_lane", 4), ("64_lane", 2)] {
        if name.ends_with(tag) {
            return Some(lanes);
        }
    }
    None
}

fn quoted_module_is_malformed(pairs: pest::iterators::Pairs<Rule>) -> bool {
    // `$id` → (param types, result types) for every `(type $id (func …))` the
    // quoted module declares. A block that gives BOTH a `(type $t)` and an
    // inline signature must repeat `$t`'s shape exactly; the spec calls a
    // mismatch malformed ("inline function type"), and without the declared
    // shape there is nothing to compare against.
    let mut func_type_shapes: HashMap<String, (Vec<String>, Vec<String>)> = HashMap::new();
    fn collect_type_shapes(
        pair: Pair<Rule>,
        out: &mut HashMap<String, (Vec<String>, Vec<String>)>,
    ) {
        if pair.as_rule() == Rule::type_field {
            let children: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
            let name = children
                .iter()
                .find(|c| c.as_rule() == Rule::id)
                .map(|c| c.as_str().trim_start_matches('$').to_string());
            if let Some(name) = name {
                let mut params = Vec::new();
                let mut results = Vec::new();
                fn scan(p: Pair<Rule>, params: &mut Vec<String>, results: &mut Vec<String>) {
                    match p.as_rule() {
                        Rule::param => params.extend(
                            p.into_inner()
                                .filter(|c| {
                                    matches!(c.as_rule(), Rule::any_val_type | Rule::val_type)
                                })
                                .map(|c| c.as_str().to_string()),
                        ),
                        Rule::result => results.extend(
                            p.into_inner()
                                .filter(|c| {
                                    matches!(c.as_rule(), Rule::any_val_type | Rule::val_type)
                                })
                                .map(|c| c.as_str().to_string()),
                        ),
                        _ => {
                            for c in p.into_inner() {
                                scan(c, params, results);
                            }
                        }
                    }
                }
                for c in pair.clone().into_inner() {
                    scan(c, &mut params, &mut results);
                }
                out.insert(name, (params, results));
            }
        }
        for c in pair.into_inner() {
            collect_type_shapes(c, out);
        }
    }
    for p in pairs.clone() {
        collect_type_shapes(p, &mut func_type_shapes);
    }
    let walk = |pair: Pair<Rule>| -> bool {
        fn walk_inner(
            pair: Pair<Rule>,
            shapes: &HashMap<String, (Vec<String>, Vec<String>)>,
        ) -> bool {
            let walk = |p: Pair<Rule>| walk_inner(p, shapes);
        if matches!(
            pair.as_rule(),
            Rule::export_inline | Rule::import_inline | Rule::export_field | Rule::import_field
        ) && pair
            .clone()
            .into_inner()
            .filter(|c| c.as_rule() == Rule::string)
            .any(|s| !name_string_is_utf8(s.as_str()))
        {
            return true;
        }
        if pair.as_rule() == Rule::plain_instr || pair.as_rule() == Rule::folded_instr {
            let mut name: Option<String> = None;
            let mut first_int: Option<String> = None;
            let mut first_float: Option<String> = None;
            for c in pair.clone().into_inner() {
                // `align=N` states the access alignment as 2^k BYTES, so N must
                // be a positive power of two. `align=0` and `align=7` are
                // malformed TEXT (`align.wast`), separately from `align=8` on a
                // one-byte access, which parses and is merely INVALID.
                if c.as_rule() == Rule::instr_arg {
                    if let Some(digits) = c.as_str().trim().strip_prefix("align=") {
                        match digits.parse::<u32>() {
                            Ok(n) if n.is_power_of_two() => {}
                            _ => return true,
                        }
                    }
                    // A memarg `offset=` field is a WAT `uN` — decimal or
                    // `0x`-hex, `_` separators allowed, NO SIGN. `offset=-1`
                    // is not a memarg the lexer can build at all, which is why
                    // the spec's expected text is "unknown operator" and not
                    // something about offsets (`simd_address.wast`).
                    //
                    // ⛔ This is a LEXICAL check only. `offset=4294967296` on a
                    // 32-bit memory lexes fine and is INVALID, not malformed —
                    // flagging it here would make that assertion pass for the
                    // wrong reason.
                    if let Some(digits) = c.as_str().trim().strip_prefix("offset=") {
                        if !is_wat_unsigned_literal(digits) {
                            return true;
                        }
                    }
                }
                match c.as_rule() {
                    Rule::instr_name if name.is_none() => name = Some(c.as_str().to_string()),
                    Rule::instr_arg if first_int.is_none() && first_float.is_none() => {
                        if let Some(inner) = c.into_inner().next() {
                            match inner.as_rule() {
                                Rule::integer => first_int = Some(inner.as_str().to_string()),
                                Rule::float => first_float = Some(inner.as_str().to_string()),
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            if let Some(n) = name.as_deref() {
                if unknown_instruction_name(n) {
                    return true;
                }
                if let Some(lit) = first_int.as_deref() {
                    if const_literal_out_of_range(n, lit) {
                        return true;
                    }
                }
                // An INTEGER const takes an integer literal. The grammar's
                // `float` rule matches `nan`, `nan:arithmetic`, `inf` and any
                // decimal-point form, and the generic operand rule lets one
                // reach `i32.const` — which the spec calls malformed
                // (`i32.wast`, `i64.wast`: "unexpected token").
                if first_float.is_some() && matches!(n, "i32.const" | "i64.const") {
                    return true;
                }
                // `nan:canonical` / `nan:arithmetic` are RESULT PATTERNS — they
                // name a set of NaNs an `assert_return` will accept, not a
                // value. As the operand of a `const` there is nothing for them
                // to denote, and the spec calls the text malformed
                // (`f32.wast`, `f64.wast`).
                if matches!(n, "f32.const" | "f64.const")
                    && first_float
                        .as_deref()
                        .is_some_and(|f| f.starts_with("nan:"))
                {
                    return true;
                }
                // A numeric `const` MUST carry its literal. Our grammar spells
                // the operand as an optional `instr_arg`, so `(i32.const)`
                // parsed and pushed a default (`const.wast`).
                if matches!(n, "i32.const" | "i64.const" | "f32.const" | "f64.const")
                    && first_int.is_none()
                    && first_float.is_none()
                {
                    return true;
                }
                // …and it must be representable at its width.
                if matches!(n, "f32.const" | "f64.const") {
                    let is_f32 = n == "f32.const";
                    if first_float
                        .as_deref()
                        .or(first_int.as_deref())
                        .is_some_and(|f| float_literal_out_of_range(f, is_f32))
                    {
                        return true;
                    }
                }
            }
        }
        // `catch` / `catch_all` are CLAUSES of a `try_table`, not instructions.
        // Our grammar's generic mnemonic rule matches them anywhere, so
        // `(func (catch_all))` parsed (`exceptions/try_table.wast`).
        if matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            let head = pair
                .clone()
                .into_inner()
                .find(|c| c.as_rule() == Rule::instr_name)
                .map(|c| c.as_str().to_string())
                .unwrap_or_default();
            if matches!(
                head.as_str(),
                "catch" | "catch_ref" | "catch_all" | "catch_all_ref"
            ) {
                return true;
            }
            // A lane INDEX immediate is a `laneidx` — one unsigned byte. A
            // negative or >255 index has no encoding, so the text is malformed
            // (an index merely past the shape's lane count is INVALID, a
            // different assertion). `simd_lane.wast`.
            if head.ends_with("extract_lane")
                || head.ends_with("extract_lane_s")
                || head.ends_with("extract_lane_u")
                || head.ends_with("replace_lane")
                || head == "i8x16.shuffle"
            {
                // The immediate is MANDATORY: `(i8x16.extract_lane_s (local.get
                // 0) …)` has a folded operand where the laneidx must be, which
                // the spec calls malformed rather than defaulting to lane 0.
                let has_immediate = pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .any(|c| {
                        c.into_inner()
                            .next()
                            .is_some_and(|i| i.as_rule() == Rule::integer)
                    });
                if !has_immediate {
                    return true;
                }
                // `shuffle` carries SIXTEEN of them; the others exactly one.
                for idx in pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .filter_map(|c| {
                        c.into_inner()
                            .next()
                            .filter(|i| i.as_rule() == Rule::integer)
                            .map(|i| i.as_str().to_string())
                    })
                {
                    // A laneidx is `u8` — an UNSIGNED literal, so even a `+`
                    // sign has no spelling (`simd_lane.wast`).
                    let text = idx.replace('_', "");
                    if text.starts_with('+') || text.starts_with('-') {
                        return true;
                    }
                    let value = match text.strip_prefix("0x") {
                        Some(hex) => u128::from_str_radix(hex, 16).ok(),
                        None => text.parse::<u128>().ok(),
                    };
                    if value.is_none_or(|v| v > 255) {
                        return true;
                    }
                    if head != "i8x16.shuffle" {
                        break;
                    }
                }
            }
            // Every lane of a `v128.const` must fit its SHAPE's lane width,
            // signed or unsigned (`simd_const.wast`).
            if head == "v128.const" {
                let mut shape: Option<String> = None;
                for arg in pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                {
                    let Some(inner) = arg.into_inner().next() else {
                        continue;
                    };
                    match inner.as_rule() {
                        Rule::bare_lane_type | Rule::bare_val_type => {
                            shape = Some(inner.as_str().to_string())
                        }
                        Rule::integer => {
                            let bits = match shape.as_deref() {
                                Some("i8x16") => 8u32,
                                Some("i16x8") => 16,
                                Some("i32x4") => 32,
                                Some("i64x2") => 64,
                                // A FLOAT shape's lane may be written without a
                                // decimal point — `f32x4 340282356779733661637539395458142568448`
                                // is a float literal spelled as an integer, and
                                // it still has to be representable.
                                Some("f32x4") | Some("f64x2") => {
                                    if float_literal_out_of_range(
                                        inner.as_str(),
                                        shape.as_deref() == Some("f32x4"),
                                    ) {
                                        return true;
                                    }
                                    continue;
                                }
                                _ => continue,
                            };
                            if lane_literal_out_of_range(inner.as_str(), bits) {
                                return true;
                            }
                        }
                        Rule::float => {
                            let is_f32 = match shape.as_deref() {
                                Some("f32x4") => true,
                                Some("f64x2") => false,
                                _ => continue,
                            };
                            if float_literal_out_of_range(inner.as_str(), is_f32) {
                                return true;
                            }
                        }
                        _ => {}
                    }
                }
            }
            // …and it must carry its SHAPE plus exactly one lane per element:
            // `(v128.const)` and `(v128.const i32x4 0 0 0)` have no encoding.
            if head == "v128.const" {
                let mut shape: Option<String> = None;
                let mut lanes = 0usize;
                for arg in pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                {
                    match arg.into_inner().next().map(|i| (i.as_rule(), i.as_str().to_string())) {
                        Some((Rule::bare_lane_type, t)) | Some((Rule::bare_val_type, t)) => {
                            shape = Some(t)
                        }
                        Some((Rule::integer, _)) | Some((Rule::float, _)) => lanes += 1,
                        _ => {}
                    }
                }
                let expected = match shape.as_deref() {
                    Some("i8x16") => 16,
                    Some("i16x8") => 8,
                    Some("i32x4") | Some("f32x4") => 4,
                    Some("i64x2") | Some("f64x2") => 2,
                    // No shape at all.
                    _ => return true,
                };
                if lanes != expected {
                    return true;
                }
            }
            // `i8x16.shuffle` takes EXACTLY 16 lane indices — the mask is one
            // byte per result lane, so a shorter or longer list has no encoding
            // (`simd_lane.wast`).
            if head == "i8x16.shuffle" {
                let lanes = pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .filter(|c| {
                        c.clone()
                            .into_inner()
                            .next()
                            .is_some_and(|i| i.as_rule() == Rule::integer)
                    })
                    .count();
                if lanes != 16 {
                    return true;
                }
            }
        }
        // A BLOCK's parameters are positional and cannot be named — only a
        // function's `(param $x t)` may carry an id (`block.wast`).
        if pair.as_rule() == Rule::block_type
            && pair.as_str().trim_start().trim_start_matches('(').trim_start().starts_with("param")
            && pair.clone().into_inner().any(|c| c.as_rule() == Rule::id)
        {
            return true;
        }
        // A block type is written `(type)? (param)* (result)*`, in that order
        // (§6.5.3). Our `block_type*` accepts any interleaving, so
        // `(block (result i32) (param i32))` parsed — the spec calls it
        // malformed (`block.wast`, `if.wast`, `loop.wast`).
        if matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
            // 0 = (type …), 1 = (param …), 2 = (result …): the rank must never
            // decrease across the run.
            let mut rank = 0u8;
            for bt in pair
                .clone()
                .into_inner()
                .filter(|c| c.as_rule() == Rule::block_type)
            {
                let text = bt.as_str().trim_start().trim_start_matches('(').trim_start();
                let this = if text.starts_with("type") {
                    0
                } else if text.starts_with("param") {
                    1
                } else {
                    2
                };
                if this < rank {
                    return true;
                }
                rank = this;
            }
        }
        // Two fields of one struct cannot share a name: `$x` would have to
        // resolve to two different indices (`gc/struct.wast`).
        if pair.as_rule() == Rule::struct_type {
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
            for field in pair
                .clone()
                .into_inner()
                .filter(|c| c.as_rule() == Rule::field_def)
            {
                if let Some(id) = field.into_inner().find(|c| c.as_rule() == Rule::id) {
                    if !seen.insert(id.as_str().to_string()) {
                        return true;
                    }
                }
            }
        }
            // A block type written as BOTH `(type $t)` and an inline signature
            // must repeat `$t`'s shape exactly (`block.wast`, `if.wast`,
            // `loop.wast`).
            if matches!(pair.as_rule(), Rule::plain_instr | Rule::folded_instr) {
                let mut named: Option<String> = None;
                let mut params: Vec<String> = Vec::new();
                let mut results: Vec<String> = Vec::new();
                let mut has_inline = false;
                for bt in pair
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::block_type)
                {
                    let text = bt.as_str().trim_start().trim_start_matches('(').trim_start();
                    if text.starts_with("type") {
                        named = bt
                            .clone()
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::index)
                            .map(|c| c.as_str().trim_start_matches('$').to_string());
                        continue;
                    }
                    has_inline = true;
                    let types: Vec<String> = bt
                        .clone()
                        .into_inner()
                        .filter(|c| matches!(c.as_rule(), Rule::any_val_type | Rule::val_type))
                        .map(|c| c.as_str().to_string())
                        .collect();
                    if text.starts_with("param") {
                        params.extend(types);
                    } else {
                        results.extend(types);
                    }
                }
                if has_inline {
                    if let Some(declared) = named.and_then(|n| shapes.get(&n)) {
                        if declared.0 != params || declared.1 != results {
                            return true;
                        }
                    }
                }
            }
            pair.into_inner().any(walk)
        }
        walk_inner(pair, &func_type_shapes)
    };
    pairs.into_iter().any(walk)
}

/// Does a quoted name literal decode to valid UTF-8?
///
/// The escapes have to be resolved to BYTES first: `\ff` is the single byte
/// 0xFF, which is not valid UTF-8, while decoding it into a `char` would give
/// U+00FF and answer the wrong question.
fn name_string_is_utf8(literal: &str) -> bool {
    let body = literal
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .unwrap_or(literal);
    let src: Vec<char> = body.chars().collect();
    let mut bytes: Vec<u8> = Vec::with_capacity(src.len());
    let mut i = 0;
    while i < src.len() {
        if src[i] != '\\' || i + 1 >= src.len() {
            let mut buf = [0u8; 4];
            bytes.extend_from_slice(src[i].encode_utf8(&mut buf).as_bytes());
            i += 1;
            continue;
        }
        match src[i + 1] {
            'n' => {
                bytes.push(b'\n');
                i += 2;
            }
            't' => {
                bytes.push(b'\t');
                i += 2;
            }
            'r' => {
                bytes.push(b'\r');
                i += 2;
            }
            '\\' => {
                bytes.push(b'\\');
                i += 2;
            }
            '\'' => {
                bytes.push(b'\'');
                i += 2;
            }
            '"' => {
                bytes.push(b'"');
                i += 2;
            }
            // `\u{...}` is a scalar value, always valid UTF-8 once encoded.
            'u' => {
                let mut j = i + 2;
                while j < src.len() && src[j] != '}' {
                    j += 1;
                }
                i = (j + 1).min(src.len());
            }
            _ => {
                // `\XX` — two hex digits, one raw byte.
                if i + 2 < src.len() {
                    let hex: String = src[i + 1..i + 3].iter().collect();
                    if let Ok(b) = u8::from_str_radix(&hex, 16) {
                        bytes.push(b);
                        i += 3;
                        continue;
                    }
                }
                i += 2;
            }
        }
    }
    std::str::from_utf8(&bytes).is_ok()
}

/// A WAT `uN`: decimal or `0x`-prefixed hex, `_` digit separators allowed, no
/// sign and at least one digit. Shared by the memarg malformed check and
/// `parse_memarg_number`, so a field this accepts is a field that one reads.
fn is_wat_unsigned_literal(text: &str) -> bool {
    let t = text.replace('_', "");
    match t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
        Some(hex) => !hex.is_empty() && hex.chars().all(|c| c.is_ascii_hexdigit()),
        None => !t.is_empty() && t.chars().all(|c| c.is_ascii_digit()),
    }
}

/// Names the grammar's generic mnemonic rule matches that are not instructions.
/// Structural keywords (`then`, `else`, `item`, …) reach the same rule, so they
/// are not evidence of malformed text.
fn unknown_instruction_name(name: &str) -> bool {
    const STRUCTURAL: &[&str] = &[
        "block", "loop", "if", "then", "else", "end", "try_table", "catch", "catch_ref",
        "catch_all", "catch_all_ref", "item", "offset", "type", "param", "result", "local",
        "func", "table", "memory", "global", "elem", "data", "export", "import", "mut", "declare",
        "quote", "binary", "module", "tag", "field", "struct", "array", "sub", "rec", "final",
        "cont",
    ];
    if STRUCTURAL.contains(&name) {
        return false;
    }
    let base = name.split_once("@@").map(|(b, _)| b).unwrap_or(name);
    vybe_runtime::opcode::Op::from_wasm_name(base).is_none()
}

/// `iN.const` with a literal that does not fit N bits. The spec's own wording
/// is "constant out of range"; the grammar cannot know the width because the
/// literal rule is shared with every other immediate.
fn const_literal_out_of_range(name: &str, lit: &str) -> bool {
    let (bits, signed_min) = match name {
        "i32.const" => (32u32, i64::from(i32::MIN)),
        "i64.const" => (64u32, i64::MIN),
        _ => return false,
    };
    let text = lit.trim();
    let (neg, digits) = match text.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, text.trim_start_matches('+')),
    };
    let magnitude = if let Some(hex) = digits.strip_prefix("0x").or(digits.strip_prefix("0X")) {
        u128::from_str_radix(&hex.replace('_', ""), 16)
    } else {
        digits.replace('_', "").parse::<u128>()
    };
    // A literal too big for u128 is beyond every width by definition.
    let Ok(magnitude) = magnitude else {
        return true;
    };
    if neg {
        magnitude > signed_min.unsigned_abs() as u128
    } else if bits == 32 {
        // Both spellings are accepted: signed down to i32::MIN, unsigned up to
        // u32::MAX.
        magnitude > u128::from(u32::MAX)
    } else {
        magnitude > u128::from(u64::MAX)
    }
}

/// A `v128.const` lane literal that fits neither the signed nor the unsigned
/// range of its shape's lane width. WAT accepts either spelling per lane, so
/// `255` and `-128` are both legal i8 lanes while `256` and `-129` are not.
fn lane_literal_out_of_range(lit: &str, bits: u32) -> bool {
    let text = lit.trim().replace('_', "");
    let negative = text.starts_with('-');
    let magnitude = text.trim_start_matches(['+', '-']);
    let value = match magnitude.strip_prefix("0x").or_else(|| magnitude.strip_prefix("0X")) {
        Some(hex) => u128::from_str_radix(hex, 16),
        None => magnitude.parse::<u128>(),
    };
    // Beyond u128 is beyond every lane width by definition.
    let Ok(value) = value else { return true };
    if negative {
        value > 1u128 << (bits - 1)
    } else {
        value >= 1u128 << bits
    }
}

/// A float literal whose magnitude has no finite representation at the target
/// width. `f32.const 0x1p128` rounds to infinity, and the spec calls the TEXT
/// malformed ("constant out of range") rather than accepting an infinity nobody
/// wrote (`const.wast`, `simd_const.wast`).
fn float_literal_out_of_range(lit: &str, is_f32: bool) -> bool {
    let lower = lit.trim().replace('_', "").to_ascii_lowercase();
    let magnitude = lower.trim_start_matches(['+', '-']);
    if magnitude.starts_with("inf") {
        return false;
    }
    if let Some(rest) = magnitude.strip_prefix("nan") {
        // A NaN PAYLOAD occupies the significand field: 23 bits for f32, 52 for
        // f64, and it must be non-zero (a zero payload is an infinity, not a
        // NaN). `nan:0x80_0000` overflows f32's field (`simd_const.wast`).
        let Some(hex) = rest.strip_prefix(":0x") else {
            return false;
        };
        let bits = if is_f32 { 23 } else { 52 };
        return match u128::from_str_radix(hex, 16) {
            Ok(payload) => payload == 0 || payload >= 1u128 << bits,
            Err(_) => true,
        };
    }
    let value = match magnitude.strip_prefix("0x") {
        Some(rest) => {
            let (mantissa, exponent) = match rest.split_once('p') {
                Some((m, e)) => (m, e.parse::<i32>().unwrap_or(0)),
                None => (rest, 0),
            };
            let (int, frac) = match mantissa.split_once('.') {
                Some((i, f)) => (i, f),
                None => (mantissa, ""),
            };
            let mut v = u128::from_str_radix(int, 16)
                .map(|n| n as f64)
                .unwrap_or(f64::INFINITY);
            let mut scale = 1.0f64 / 16.0;
            for c in frac.chars() {
                v += f64::from(c.to_digit(16).unwrap_or(0)) * scale;
                scale /= 16.0;
            }
            v * 2f64.powi(exponent)
        }
        None => magnitude.parse::<f64>().unwrap_or(f64::INFINITY),
    };
    if !value.is_finite() {
        return true;
    }
    is_f32 && (value as f32).is_infinite()
}

/// A run of digits with underscores only BETWEEN digits, and never empty.
///
/// `_100`, `99_`, `1__000`, `_0x100` and the empty body of `0x` are all
/// malformed number tokens (`int_literals.wast`, `const.wast`), and our
/// `integer` grammar rule accepts every one of them because `_` is simply
/// listed as a digit character.
fn wat_digits_ok(s: &str, hex: bool) -> bool {
    let mut pending = true; // no digit yet — a leading `_` is invalid
    for c in s.chars() {
        if c == '_' {
            if pending {
                return false;
            }
            pending = true;
        } else if if hex {
            c.is_ascii_hexdigit()
        } else {
            c.is_ascii_digit()
        } {
            pending = false;
        } else {
            return false;
        }
    }
    !pending
}

/// Is this idchar run a well-formed WAT number literal (integer or float)?
fn wat_number_token_ok(tok: &str) -> bool {
    let body = tok.strip_prefix(['+', '-']).unwrap_or(tok);
    if body == "inf" {
        return true;
    }
    if let Some(rest) = body.strip_prefix("nan") {
        if rest.is_empty() {
            return true;
        }
        if let Some(hex) = rest.strip_prefix(":0x") {
            return wat_digits_ok(hex, true);
        }
        return matches!(rest, ":canonical" | ":arithmetic");
    }
    // `mantissa (exponent)?`, split on the exponent marker for the base.
    fn split_exp(s: &str, marker: [char; 2]) -> (&str, Option<&str>) {
        match s.split_once(marker) {
            Some((m, e)) => (m, Some(e)),
            None => (s, None),
        }
    }
    let (hex, digits) = match body.strip_prefix("0x").or_else(|| body.strip_prefix("0X")) {
        Some(rest) => (true, rest),
        None => (false, body),
    };
    let (mantissa, exponent) = split_exp(digits, if hex { ['p', 'P'] } else { ['e', 'E'] });
    let mantissa_ok = match mantissa.split_once('.') {
        // A fraction may be empty (`1.`), an integer part may NOT (`.5`).
        Some((int, frac)) => {
            wat_digits_ok(int, hex) && (frac.is_empty() || wat_digits_ok(frac, hex))
        }
        None => wat_digits_ok(mantissa, hex),
    };
    // The exponent is always DECIMAL, even for a hex float.
    let exponent_ok = exponent.is_none_or(|e| wat_digits_ok(e.strip_prefix(['+', '-']).unwrap_or(e), false));
    mantissa_ok && exponent_ok
}

/// Text-level malformities that only a LEXER can see, asked of `assert_malformed`
/// quoted source alone.
///
/// WAT lexes a maximal run of idchars as ONE token. Our grammar instead matches
/// an instruction name against a list and a number with `_` treated as a digit,
/// so `i32.const0` parses as `i32.const` then `0`, `i32.load32` as `i32.load`
/// then `32`, and `_100` as a number — the token SEPARATION the spec requires is
/// simply not represented. Rather than tighten rules every well-formed file also
/// goes through, the maximal runs are re-derived here and checked.
///
/// Deliberately conservative: an unknown dotted mnemonic is reported only when
/// trimming a trailing digit (or `_s`/`_u` then digits) run turns it back into a
/// KNOWN one, which is exactly the glued-token shape. Reporting more would make
/// an `assert_malformed` pass for the wrong reason, and a wrongly-passing
/// malformed assertion hides real leniency.
/// `block … end $l` — the id after `end` (or after `else`) must be the id the
/// block was OPENED with; the spec calls a disagreement "mismatching label" and
/// classifies it as MALFORMED, not invalid. An unnamed block mismatches every
/// id, which is `(func block end $l)` in `block.wast`, `loop.wast` and
/// `if.wast`.
///
/// This is a property of how the two tokens PAIR UP, so it is checked on the
/// token stream rather than on the parse tree: the grammar sees `end` and a
/// following id as an instruction and its argument and has nowhere to put the
/// relation. Only the PLAIN form can be wrong — a folded `(block $l …)` is
/// closed by its paren and never writes an id — so an opener preceded by `(`
/// is skipped.
fn quoted_text_has_label_mismatch(src: &str) -> bool {
    #[derive(PartialEq)]
    enum Tok {
        Open,
        Close,
        Word(String),
    }
    // Reuse the same idchar notion as the token scanner; strings and comments
    // are skipped whole so a `;; end $x` cannot open or close anything.
    const IDCHAR_EXTRA: &str = "!#$%&'*+-./:<=>?@\\^_`|~";
    let is_idchar = |c: char| c.is_ascii_alphanumeric() || IDCHAR_EXTRA.contains(c);
    let chars: Vec<char> = src.chars().collect();
    let mut toks: Vec<Tok> = Vec::new();
    let mut i = 0usize;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            i += 1;
            while i < chars.len() {
                if chars[i] == '\\' {
                    i += 2;
                    continue;
                }
                if chars[i] == '"' {
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == ';' && chars.get(i + 1) == Some(&';') {
            while i < chars.len() && chars[i] != '\n' {
                i += 1;
            }
            continue;
        }
        if c == '(' && chars.get(i + 1) == Some(&';') {
            i += 2;
            while i + 1 < chars.len() && !(chars[i] == ';' && chars[i + 1] == ')') {
                i += 1;
            }
            i += 2;
            continue;
        }
        if c == '(' {
            toks.push(Tok::Open);
            i += 1;
            continue;
        }
        if c == ')' {
            toks.push(Tok::Close);
            i += 1;
            continue;
        }
        if !is_idchar(c) {
            i += 1;
            continue;
        }
        let start = i;
        while i < chars.len() && is_idchar(chars[i]) {
            i += 1;
        }
        toks.push(Tok::Word(chars[start..i].iter().collect()));
    }
    // Each entry is the label a still-open PLAIN block was written with.
    let mut open: Vec<Option<String>> = Vec::new();
    for (n, t) in toks.iter().enumerate() {
        let Tok::Word(w) = t else { continue };
        let prev_is_open = n > 0 && toks[n - 1] == Tok::Open;
        let next_id = match toks.get(n + 1) {
            Some(Tok::Word(x)) if x.starts_with('$') => Some(x.clone()),
            _ => None,
        };
        match w.as_str() {
            "block" | "loop" | "if" | "try_table" if !prev_is_open => open.push(next_id),
            // `else` names the SAME label as its `if`; the block stays open.
            "else" if !prev_is_open => {
                if let Some(id) = next_id {
                    if open.last().cloned().flatten() != Some(id) {
                        return true;
                    }
                }
            }
            "end" => {
                if let Some(id) = next_id {
                    if open.last().cloned().flatten() != Some(id) {
                        return true;
                    }
                }
                open.pop();
            }
            _ => {}
        }
    }
    false
}

fn quoted_text_has_bad_token(src: &str) -> bool {
    const IDCHAR_EXTRA: &str = "!#$%&'*+-./:<=>?@\\^_`|~";
    let is_idchar = |c: char| c.is_ascii_alphanumeric() || IDCHAR_EXTRA.contains(c);
    let chars: Vec<char> = src.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        let c = chars[i];
        // Strings are not token runs; skip them whole, honouring `\"`.
        if c == '"' {
            i += 1;
            while i < chars.len() {
                if chars[i] == '\\' {
                    i += 2;
                    continue;
                }
                if chars[i] == '"' {
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == ';' && chars.get(i + 1) == Some(&';') {
            while i < chars.len() && chars[i] != '\n' {
                i += 1;
            }
            continue;
        }
        if c == '(' && chars.get(i + 1) == Some(&';') {
            i += 2;
            while i + 1 < chars.len() && !(chars[i] == ';' && chars[i + 1] == ')') {
                i += 1;
            }
            i += 2;
            continue;
        }
        if !is_idchar(c) {
            i += 1;
            continue;
        }
        let start = i;
        while i < chars.len() && is_idchar(chars[i]) {
            i += 1;
        }
        let token: String = chars[start..i].iter().collect();

        // A lone `$` names nothing. WASM 3.0's QUOTED identifier is `$"name"`
        // with no space between the two — `$ "a"` is an empty id followed by a
        // string, which `id.wast` calls malformed for exactly that reason.
        if token == "$" {
            if chars.get(i) != Some(&'"') {
                return true;
            }
            let mut j = i + 1;
            let mut body = String::new();
            while j < chars.len() && chars[j] != '"' {
                if chars[j] == '\\' && j + 1 < chars.len() {
                    body.push(chars[j]);
                    body.push(chars[j + 1]);
                    j += 2;
                    continue;
                }
                body.push(chars[j]);
                j += 1;
            }
            // Empty, containing a raw control character, or not valid UTF-8
            // once its escapes are resolved to BYTES.
            if body.is_empty()
                || body.chars().any(|c| c.is_control())
                || !name_string_is_utf8(&format!("\"{body}\""))
            {
                return true;
            }
            i = j + 1;
            continue;
        }
        if token.starts_with('$') {
            continue;
        }
        let head = token.chars().next().unwrap_or(' ');
        if head.is_ascii_digit() || head == '+' || head == '-' || head == '.' {
            if !wat_number_token_ok(&token) {
                return true;
            }
            continue;
        }
        // A WAT keyword starts with a LOWERCASE LETTER (§6.2), an identifier
        // with `$`, a number with a digit or sign. A run starting with `_` or
        // an uppercase letter is none of those — it is a RESERVED token, and
        // `_100` is not "100 with a decoration", it is not a number at all
        // (`int_literals.wast`). `@`, `!` and the other idchars stay allowed:
        // the annotations proposal spells its ids with them.
        if head == '_' || head.is_ascii_uppercase() {
            return true;
        }
        if unknown_instruction_name(&token) {
            let core = token
                .trim_end_matches(['s', 'u'])
                .trim_end_matches('_')
                .trim_end_matches(|c: char| c.is_ascii_digit());
            if core != token && !core.is_empty() && !unknown_instruction_name(core) {
                return true;
            }
        }
    }
    false
}

/// `(assert_invalid (module …) "diagnostic")` — the module must PARSE and then
/// fail validation. Settled at walk time, like `assert_malformed`: validity is
/// a static property of the text, so nothing needs to run.
fn walk_assert_invalid(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let _ = &__w;
    let span = to_span(&pair);
    // Both spellings: an inline `(module …)` is already parsed, a `(module
    // quote "…")` has to be parsed from its concatenated strings first.
    let mut reason: Option<String> = None;
    let mut examined = false;
    // ⛔ The expected diagnostic is written AFTER the module, and the module
    // arm below `break`s — so reading both in one pass silently never saw the
    // string and every comparison came out vacuously true. Collect first.
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let expected: Option<String> = children
        .iter()
        .find(|c| c.as_rule() == Rule::string)
        .map(|c| unquote(c.as_str()));
    for child in children {
        // The expected diagnostic. ⛔ It used to be parsed and DROPPED, which
        // made every `assert_invalid` assert only that the module was rejected
        // for SOME reason — a module refused for a bad lane index satisfied a
        // fixture demanding "alignment must not be larger than natural". That
        // is the identical disease `walk_assert_trap` records against its own
        // past, and it is how an over-flagging check makes an assertion pass
        // for the WRONG reason while looking green.
        if child.as_rule() != Rule::module_or_binary {
            continue;
        }
        let text = child.as_str();
        let head: Vec<&str> = text.split_whitespace().take(4).collect();
        if head.iter().any(|t| *t == "binary") {
            // A binary module's validity is a property of its BYTES; the text
            // walk here cannot see them. Left unexamined on purpose.
            break;
        }
        if head.iter().any(|t| *t == "quote") {
            let mut source = String::new();
            for sp in child.clone().into_inner() {
                if sp.as_rule() == Rule::string {
                    source.push_str(&unquote(sp.as_str()));
                }
            }
            // ⛔ A QUOTED MODULE MAY OMIT ITS OWN `(module …)` WRAPPER. Both
            // spellings are in the suite: align.wast quotes
            // `"(module (memory 0) …)"` while address.wast quotes `"(memory 1)"`
            // and `"(func …)"` — the text format's inline-module form. Parsed
            // as-is the second yields module FIELDS at top level and no
            // `Rule::module` pair at all, so every module-level rule (name
            // resolution, limits, memarg range, stack typing) silently never
            // ran and the assertion reported "the module validated".
            let trimmed = source.trim_start();
            let wrapped;
            let text = if trimmed.starts_with("(module") {
                &source
            } else {
                wrapped = format!("(module {source})");
                &wrapped
            };
            if let Ok(pairs) = WastParser::parse(Rule::program, text) {
                examined = true;
                reason = module_invalid_reason(pairs);
            }
            break;
        }
        examined = true;
        let mut names = std::collections::HashSet::new();
        reason = module_invalid_walk(child.clone(), &mut names);
        if reason.is_none() {
            reason = stack_typing_reason_in(&child);
        }
        break;
    }
    // ⛔ ONE DIRECTION ONLY: the reason we give must CONTAIN the asserted text,
    // which is the reference harness's own rule. Accepting the reverse as well
    // looked like harmless latitude and was not — a fixture asserting
    // "unknown memory 1" was satisfied by a bare "unknown memory", so a check
    // that cannot name the index discharged an assertion about the index. Being
    // vaguer than the fixture is exactly the failure this comparison exists to
    // catch, and it is the direction that flatters us.
    let matched = match (&reason, &expected) {
        (Some(r), Some(e)) => r.contains(e.as_str()),
        (Some(_), None) => true,
        _ => false,
    };
    Ok(match (&reason, matched) {
        // Rejected for the reason asserted: discharged.
        (Some(_), true) => Statement::with_span(StmtKind::Empty, span),
        (Some(r), false) => Statement::with_span(
            StmtKind::Throw {
                expr: Some(Expression::string(&format!(
                    "assert_invalid failed: expected \"{}\", got \"{}\"",
                    expected.as_deref().unwrap_or(""),
                    r
                ))),
                cause: None,
            },
            span,
        ),
        _ => Statement::with_span(
            StmtKind::Throw {
                expr: Some(Expression::string(&format!(
                    "assert_invalid failed: {} — expected \"{}\"",
                    if examined {
                        "the module validated"
                    } else {
                        "the module was not examined"
                    },
                    expected.as_deref().unwrap_or("")
                ))),
                cause: None,
            },
            span,
        ),
    })
}

fn walk_assert_malformed(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut quoted: Option<String> = None;
    let mut binary_bytes: Option<Vec<u8>> = None;
    for child in pair.into_inner() {
        if child.as_rule() != Rule::module_or_binary {
            continue;
        }
        let text = child.as_str();
        // `(module $id? binary "…")` vs `(module $id? quote "…")` — the
        // keyword sits after an optional id, so match it as a token rather
        // than by a prefix test. `take(4)` because `definition` may sit
        // between `module` and the keyword: `(module definition binary "…")`.
        let head: Vec<&str> = text.split_whitespace().take(4).collect();
        if head.iter().any(|t| *t == "binary") {
            binary_bytes = Some(binary_module_bytes(&child));
            break;
        }
        if !head.iter().any(|t| *t == "quote") {
            break;
        }
        // The module text is the concatenation of every string literal in the
        // `quote` form — the spec splits long fixtures across several.
        let mut source = String::new();
        for s in child.into_inner() {
            if s.as_rule() == Rule::string {
                source.push_str(&unquote(s.as_str()));
            }
        }
        quoted = Some(source);
        break;
    }
    if let Some(bytes) = binary_bytes {
        // ⛔ THIS USED TO BE `Empty` — an assertion that examined NOTHING.
        //
        // `assert_malformed` on a `(module binary …)` asserts that those BYTES
        // do not decode. Discharging it without decoding them reports PASS for
        // a fixture that was never looked at, which is the same disease as a
        // checker that cannot parse its input reporting `0 problems`.
        //
        // The bytes are the concatenation of the string literals, and
        // `read_wasm` is the decoder the runtime already uses for a real
        // `.wasm`. Decoding is a STATIC property of the fixture, so it is
        // settled here at walk time exactly as the `quote` path settles
        // parsing below.
        return Ok(match vybe_platform_wasm::read_wasm(&bytes) {
            // Rejected, as the spec requires: discharged.
            Err(_) => Statement::with_span(StmtKind::Empty, span),
            // It decoded. That is the assertion failing, and it must say so.
            Ok(_) => Statement::with_span(
                StmtKind::Throw {
                    expr: Some(Expression::string(
                        "assert_malformed failed: the binary module decoded",
                    )),
                    cause: None,
                },
                span,
            ),
        });
    }
    let Some(source) = quoted else {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    };
    let source = {
        let trimmed = source.trim();
        if trimmed.starts_with("(module") {
            trimmed.to_string()
        } else {
            format!("(module {trimmed})")
        }
    };
    // The LEXICAL layer first: token separation and number-literal shape are
    // properties of the text that the grammar, matching names from a list and
    // treating `_` as a digit, cannot express.
    if quoted_text_has_bad_token(&source) || quoted_text_has_label_mismatch(&source) {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }
    let parsed = match WastParser::parse(Rule::program, &source) {
        // Rejected by the grammar, as the spec requires: discharged.
        Err(_) => return Ok(Statement::with_span(StmtKind::Empty, span)),
        Ok(pairs) => pairs,
    };
    // Two text-level malformities the grammar deliberately cannot express, so
    // they are checked over the parse tree instead of by tightening rules that
    // every well-formed file also goes through. Both are confined to this
    // assertion: nothing here runs during ordinary compilation.
    if quoted_module_is_malformed(parsed.clone()) {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }
    fn any_module_invalid(pair: Pair<Rule>) -> bool {
        if pair.as_rule() == Rule::module && validate_module(&pair).is_err() {
            return true;
        }
        pair.into_inner().any(any_module_invalid)
    }
    if parsed.into_iter().any(any_module_invalid) {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }
    Ok(Statement::with_span(
        StmtKind::Throw {
            expr: Some(Expression::string(&format!(
                "assert_malformed failed: the module parsed — {}",
                source.trim()
            ))),
            cause: None,
        },
        span,
    ))
}

fn walk_assert_exception(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    for child in pair.into_inner() {
        if child.as_rule() == Rule::action {
            action_expr = Some(walk_action(__w, child)?);
        }
    }
    let Some(action) = action_expr else {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    };
    let flag = fresh_result_temp(__w);
    let decl = Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(flag.clone()),
                type_hint: None,
                init: Some(Expression::bool(false)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        },
        span,
    );
    let try_table = Statement::with_span(
        StmtKind::WasmTryTable {
            body: vec![
                Statement::with_span(StmtKind::Expr(action), span),
                Statement::with_span(
                    StmtKind::Assign {
                        targets: vec![Expression::ident(&flag)],
                        value: Expression::bool(true),
                        by_ref: false,
                    },
                    span,
                ),
            ],
            catches: vec![WasmCatch {
                tag: None,
                payload_binds: Vec::new(),
                capture_ref: false,
                exnref_bind: None,
                body: Vec::new(),
            }],
            params: 0,
            results: 0,
        },
        span,
    );
    let check = Statement::with_span(
        StmtKind::If {
            cond: Expression::ident(&flag),
            then_body: vec![Statement::with_span(
                StmtKind::Throw {
                    expr: Some(Expression::string(
                        "assert_exception failed: expected an exception",
                    )),
                    cause: None,
                },
                span,
            )],
            elifs: Vec::new(),
            else_body: None,
        },
        span,
    );
    Ok(Statement::with_span(
        StmtKind::Block(vec![decl, try_table, check]),
        span,
    ))
}

/// `__wast_check_trap(err, expected, marker)` — the WHOLE of an `assert_trap`
/// check, emitted ONCE per script and called with three constants per
/// assertion.
///
/// It is one function rather than an inlined check for a concrete reason:
/// the per-assertion form allocated scratch locals in every catch body, and
/// `Chunk::alloc_scratch` indexes a `u16` that is only `max`ed at the END of a
/// chunk. `bulk-memory/table_copy.wast` has 1206 `assert_trap`s, which pushed
/// the counter past 65535 and panicked the compiler ("attempt to add with
/// overflow"). Collapsing the body into a call keeps each assertion at a fixed,
/// tiny cost. (The underlying `alloc_scratch` ceiling is a separate issue and
/// is not fixed here — this only stops the wast front end from being the thing
/// that reaches it.)
///
/// Three facts drive the body:
///
///  * **Two shapes of caught value.** A WASM trap is `make_runtime_error(msg)`
///    — an OBJECT carrying `.message` (`calls.rs::raise_trap`). A host-side
///    trap is a bare STRING (`ctx.throw_value(Value::String)`, e.g. every
///    `wasm:js-string` TypeError), as is our own "expected a trap" marker. So
///    read `.message` and fall back to the value itself when that is null;
///    reading `.message` off a string yields null rather than trapping, which
///    is what makes the fallback safe.
///  * **Containment, not equality.** The spec suite's convention is that the
///    fixture's text appears IN the implementation's: `"cast failure"` is
///    satisfied by `"trap: descriptor cast failure"`. Substring search is
///    composed from the `wasm:js-string` builtins the runtime already
///    registers (`length`/`substring`/`equals`) rather than a new host
///    function, which would have been a runtime change.
///  * **An empty expected message passes on any trap**, which is what the
///    spec means by an `assert_trap` written without text.
fn build_trap_check_helper(span: Span) -> Statement {
    fn param(name: &str) -> Param {
        Param {
            name: name.to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }
    }
    fn let_of(name: &str, init: Expression, span: Span) -> Statement {
        Statement::with_span(
            StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.to_string()),
                    type_hint: None,
                    init: Some(init),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            },
            span,
        )
    }
    fn bin(op: BinOp, l: Expression, r: Expression, span: Span) -> Expression {
        Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(l),
                right: Box::new(r),
            },
            span,
        )
    }
    fn if_then(cond: Expression, then_body: Vec<Statement>, span: Span) -> Statement {
        Statement::with_span(
            StmtKind::If {
                cond,
                then_body,
                elifs: Vec::new(),
                else_body: None,
            },
            span,
        )
    }
    fn ret(span: Span) -> Statement {
        Statement::with_span(StmtKind::Return(None), span)
    }

    let e = || Expression::ident("__wct_e");
    let exp = || Expression::ident("__wct_exp");
    let mk = || Expression::ident("__wct_marker");
    let m = || Expression::ident("__wct_m");
    let hl = || Expression::ident("__wct_hl");
    let nl = || Expression::ident("__wct_nl");
    let i = || Expression::ident("__wct_i");

    // throw "assert_trap failed: expected trap message containing <exp>, got: <m>"
    let report = Statement::with_span(
        StmtKind::Throw {
            expr: Some(bin(
                BinOp::Add,
                bin(
                    BinOp::Add,
                    bin(
                        BinOp::Add,
                        Expression::string("assert_trap failed: expected trap message containing \""),
                        exp(),
                        span,
                    ),
                    Expression::string("\", got: "),
                    span,
                ),
                m(),
                span,
            )),
            cause: None,
        },
        span,
    );

    let body = vec![
        // The action completed normally — our own marker came back, so no trap
        // happened at all. Re-raise it as the failure it is.
        if_then(
            bin(BinOp::Eq, e(), mk(), span),
            vec![Statement::with_span(
                StmtKind::Throw {
                    expr: Some(e()),
                    cause: None,
                },
                span,
            )],
            span,
        ),
        let_of(
            "__wct_m",
            Expression::with_span(
                ExprKind::Member {
                    object: Box::new(e()),
                    field: "message".to_string(),
                    null_safe: false,
                },
                span,
            ),
            span,
        ),
        if_then(
            bin(BinOp::Eq, m(), Expression::null(), span),
            vec![Statement::with_span(
                StmtKind::Assign {
                    targets: vec![m()],
                    value: e(),
                    by_ref: false,
                },
                span,
            )],
            span,
        ),
        let_of("__wct_nl", make_call("string_length", vec![exp()], span), span),
        // No expected text: any trap satisfies the assertion.
        if_then(
            bin(BinOp::Eq, nl(), Expression::int(0), span),
            vec![ret(span)],
            span,
        ),
        let_of("__wct_hl", make_call("string_length", vec![m()], span), span),
        // Needle longer than haystack: no window to test, and `hl - nl` would
        // go negative — `substring` clamps u32-wise, so the loop must not run.
        if_then(bin(BinOp::Lt, hl(), nl(), span), vec![report.clone()], span),
        let_of("__wct_i", Expression::int(0), span),
        Statement::with_span(
            StmtKind::While {
                cond: bin(BinOp::LtEq, i(), bin(BinOp::Sub, hl(), nl(), span), span),
                body: vec![
                    if_then(
                        bin(
                            BinOp::Eq,
                            make_call(
                                "string_equals",
                                vec![
                                    make_call(
                                        "string_substring",
                                        vec![m(), i(), bin(BinOp::Add, i(), nl(), span)],
                                        span,
                                    ),
                                    exp(),
                                ],
                                span,
                            ),
                            Expression::int(1),
                            span,
                        ),
                        vec![ret(span)],
                        span,
                    ),
                    Statement::with_span(
                        StmtKind::Assign {
                            targets: vec![i()],
                            value: bin(BinOp::Add, i(), Expression::int(1), span),
                            by_ref: false,
                        },
                        span,
                    ),
                ],
                else_body: None,
            },
            span,
        ),
        report,
    ];

    Statement::with_span(
        StmtKind::FunctionDecl {
            name: "__wast_check_trap".to_string(),
            params: vec![param("__wct_e"), param("__wct_exp"), param("__wct_marker")],
            return_type: None,
            body,
            modifiers: Default::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span,
    )
}

fn walk_assert_trap(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut action_expr: Option<Expression> = None;
    let mut expected_msg: Option<String> = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::action => action_expr = Some(walk_action(__w, child)?),
            // The expected trap text. This used to be parsed and DROPPED,
            // which made every `assert_trap` assert only that SOMETHING went
            // wrong — a `call_indirect` reporting "null is not callable"
            // satisfied a fixture demanding "uninitialized element", and the
            // 16 "cast failure" asserts in the descriptor suite passed on any
            // trap at all.
            Rule::string => expected_msg = Some(unquote(child.as_str())),
            _ => {}
        }
    }
    let Some(action) = action_expr else {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    };
    // Language-level lowering, same principle as assert_return: run the
    // action inside a try; COMPLETING normally is the failure, so the body
    // throws a marker after the action, and the catch re-raises exactly
    // that marker (a genuine trap lands in the catch and passes).
    let marker = "assert_trap failed: expected a trap";
    let body = vec![
        Statement::with_span(StmtKind::Expr(action), span),
        Statement::with_span(
            StmtKind::Throw {
                expr: Some(Expression::string(marker)),
                cause: None,
            },
            span,
        ),
    ];
    // The whole check is ONE call to the helper `parse` prepends. Inlining it
    // here instead cost scratch locals per assertion and overflowed
    // `alloc_scratch`'s u16 on the big generated fixtures — see
    // `build_trap_check_helper`.
    //
    // The marker is passed in rather than compared here because the helper has
    // to tell two cases apart that both arrive as a caught value: our own
    // "the action completed, so no trap happened" marker, and a real trap.
    let catch_body = vec![Statement::with_span(
        StmtKind::Expr(make_call(
            "__wast_check_trap",
            vec![
                Expression::ident("__wast_trap_err"),
                Expression::string(expected_msg.as_deref().unwrap_or("")),
                Expression::string(marker),
            ],
            span,
        )),
        span,
    )];
    __w.needs_trap_contains = true;
    Ok(Statement::with_span(
        StmtKind::Try {
            body,
            catches: vec![CatchClause {
                types: Vec::new(),
                var_name: Some("__wast_trap_err".to_string()),
                stack_var: None,
                body: catch_body,
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        },
        span,
    ))
}

#[allow(dead_code)]
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

fn walk_register_cmd(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::string {
            name = unquote(child.as_str());
            break;
        }
    }
    // `(register "m")` names the module just walked, which is when its exported
    // TAGS become importable as `("m", export)`. A later module's
    // `(tag $x (import "m" "e"))` resolves to that same entity.
    let exports = std::mem::take(&mut __w.pending_tag_exports);
    for (export, canon) in exports {
        __w.registered_tags.insert((name.clone(), export), canon);
    }
    // …and its FUNCTIONS: the module just walked becomes reachable under `name`,
    // so a later `(import "name" "e")` can resolve to its class.
    let class = __w.module_class_name.clone();
    __w.registered_module_class.insert(name.clone(), class);
    Ok(Statement::with_span(
        StmtKind::Expr(make_call(
            "__wasm_register",
            vec![Expression::string(&name)],
            span,
        )),
        span,
    ))
}

fn walk_get_cmd(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut module_id: Option<String> = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            // `(get $M "e")` reads the global exported by the module NAMED $M.
            Rule::id => module_id = Some(child.as_str()[1..].to_string()),
            Rule::string => {
                if name.is_empty() {
                    name = unquote(child.as_str());
                }
            }
            _ => {}
        }
    }
    // An exported global IS a top-level binding — read it directly.
    let binding = match &module_id {
        Some(m) => {
            __w.module_global_exports.get(m).and_then(|e| e.get(&name).cloned())
        }
        None => __w.export_global_map.get(&name).cloned(),
    };
    let expr = match binding {
        Some(b) => Expression::with_span(ExprKind::Ident(b), span),
        // Unresolved: keep the (profile no-op) call rather than inventing a
        // binding, so the failure stays visible instead of reading a stray name.
        None => make_call("__wasm_get", vec![Expression::string(&name)], span),
    };
    Ok(Statement::with_span(StmtKind::Expr(expr), span))
}

fn walk_action(__w: &mut WastWalker, pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("Empty action")?;
    match inner.as_rule() {
        Rule::invoke_cmd => {
            let stmt = walk_invoke_cmd(__w, inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()),
            }
        }
        Rule::get_cmd => {
            let stmt = walk_get_cmd(__w, inner)?;
            match stmt.kind {
                StmtKind::Expr(e) => Ok(e),
                _ => Ok(Expression::null()),
            }
        }
        _ => Ok(Expression::null()),
    }
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
        // Lanes may be floats as well as integers (`f32x4 1.5 nan:canonical …`).
        // Walk the children IN ORDER so lane positions are preserved — the two
        // rules cannot be collected separately and concatenated.
        for c in &children {
            match c.as_rule() {
                Rule::integer => args.push(parse_integer(c.as_str())),
                Rule::float => args.push(parse_float(c.as_str())),
                _ => {}
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
            Rule::integer => {
                // A FLOAT constant written in integer form —
                // `(f64.const 9223372036854775808)`, which the spec's own
                // conversions.wast:246 uses. There was no `f64.const` arm
                // below, so it fell through to the integer path: 2^63
                // overflows i64, `parse_integer` wrapped it negative, and the
                // conversion then trapped on a legal value. The same text with
                // a decimal point parsed correctly, which is what made this
                // invisible.
                //
                // Hex is read as a bit-pattern magnitude and widened, matching
                // `parse_integer`'s reason for going through u64: the literal
                // denotes a value, and f64 carries it exactly up to 2^53 and
                // approximately beyond — which is what the spec asks for.
                if raw_text.contains("f32.const") || raw_text.contains("f64.const") {
                    let t = child.as_str().trim().replace('_', "");
                    let lower = t.to_ascii_lowercase();
                    let stripped = lower.trim_start_matches(['+', '-']);
                    let val = if let Some(digits) = stripped.strip_prefix("0x") {
                        let mag = u128::from_str_radix(digits, 16).unwrap_or(0) as f64;
                        if lower.starts_with('-') { -mag } else { mag }
                    } else {
                        t.parse::<f64>().unwrap_or(0.0)
                    };
                    let lit = Expression::float(val);
                    return Ok(if raw_text.contains("f32.const") {
                        make_call("f32_demote_f64", vec![lit], span)
                    } else {
                        lit
                    });
                }
                let v = parse_integer(child.as_str());
                // Mirror the ACTION side's lowerings so both sides of the
                // compare carry the same value shape: `i64.const` is a BigInt
                // (a plain Int literal compiles to f64, losing bits past 2^53
                // AND the i64 result's BigInt identity); `f32.const` demotes
                // to single precision (`Value::F32`).
                if raw_text.contains("i32.const") {
                    // An `i32.const` literal is a 32-BIT PATTERN, not a number:
                    // `0x80000000` and `0xffffffff` denote i32::MIN and -1. Read
                    // at 64-bit width they stay positive and never equal the
                    // wrapped i32 the VM computes, so truncate to two's
                    // complement 32 (i32, conversions, endianness, int_exprs,
                    // memory_grow all turn on exactly this).
                    if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                        return Ok(Expression::with_span(
                            ExprKind::Lit(Literal::Int(*n as u32 as i32 as i64)),
                            span,
                        ));
                    }
                } else if raw_text.contains("i64.const") {
                    if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                        return Ok(Expression::with_span(
                            ExprKind::Lit(Literal::BigInt(*n)),
                            span,
                        ));
                    }
                } else if raw_text.contains("f32.const") {
                    return Ok(make_call("f32_demote_f64", vec![v], span));
                }
                return Ok(v);
            }
            Rule::float => {
                let v = parse_float(child.as_str());
                if raw_text.contains("f32.const") {
                    return Ok(f32_const_expr(v, span));
                }
                return Ok(v);
            }
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
            _ => unreachable!(),
        })
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
        return match s.parse::<i64>().or_else(|_| s.parse::<u64>().map(|u| u as i64)) {
            Ok(v) => Expression::int(v),
            // Too large for 64 bits, so it cannot be an integer const at all —
            // it is a FLOAT written in integer form. See `too_wide_for_integer`.
            Err(_) => Expression::float(s.parse::<f64>().unwrap_or(0.0)),
        };
    };
    // Hex: a 64-bit pattern like 0x8000000000000000 exceeds i64::MAX, so parse as
    // u64 and reinterpret to the signed value it denotes.
    match i64::from_str_radix(digits, 16)
        .or_else(|_| u64::from_str_radix(digits, 16).map(|u| u as i64))
    {
        Ok(v) => Expression::int(if neg { v.wrapping_neg() } else { v }),
        Err(_) => {
            let mut f = 0.0f64;
            for c in digits.chars() {
                f = f * 16.0 + f64::from(c.to_digit(16).unwrap_or(0));
            }
            Expression::float(if neg { -f } else { f })
        }
    }
}

/// Build the f64 a WAT NaN literal denotes, carrying BOTH its sign and its
/// payload.
///
/// `-nan` and `nan:0x…` are not decoration: `f32.reinterpret_f32`/`copysign`
/// observe the sign bit and the payload directly, so collapsing every NaN form
/// to `f64::NAN` (positive, canonical) silently rewrote the operand. That is
/// what made `conversions.wast:647` — `(f32.const -nan:0x7fffff)` — unreachable.
///
/// The payload is placed in the f64 mantissa as written. `f32_const_expr`
/// re-reads it as a 23-bit f32 payload, so the two agree without this function
/// needing to know which width it is serving. A bare `nan` (and the
/// `nan:canonical` / `nan:arithmetic` expectation forms, which pin no payload)
/// becomes the canonical quiet NaN of each width.
fn nan_literal(s: &str) -> f64 {
    let neg = s.starts_with('-');
    let payload = s
        .split_once("nan:0x")
        .and_then(|(_, hex)| u64::from_str_radix(hex.trim(), 16).ok())
        .filter(|p| *p != 0)
        .unwrap_or(0x0008_0000_0000_0000); // canonical f64 quiet bit
    f64::from_bits((u64::from(neg) << 63) | 0x7ff0_0000_0000_0000 | (payload & 0x000f_ffff_ffff_ffff))
}

/// Lower an `f32.const` operand.
///
/// The ordinary path demotes the exact-text f64 to single precision. A NaN
/// CANNOT go that way: narrowing f64→f32 is a hardware conversion, and on x86
/// `cvtsd2ss` QUIETS a signalling NaN, so `nan:0x200000` would arrive as
/// `nan:0x400000` — a payload the source never wrote. `f32.reinterpret_i32` is
/// a pure bit copy, so the pattern lands exactly as written.
fn f32_const_expr(v: Expression, span: Span) -> Expression {
    if let ExprKind::Lit(Literal::Float(f)) = &v.kind
        && f.is_nan()
    {
        let bits = f.to_bits();
        let payload = (bits & 0x000f_ffff_ffff_ffff) as u32 & 0x007f_ffff;
        // A canonical f64 NaN denotes the canonical f32 NaN, not payload 0 —
        // an all-zero mantissa with a full exponent is INFINITY, not a NaN.
        let payload = if payload == 0 { 0x0040_0000 } else { payload };
        let f32_bits = ((bits >> 32) as u32 & 0x8000_0000) | 0x7f80_0000 | payload;
        return make_call(
            "f32_reinterpret_i32",
            vec![Expression::with_span(
                ExprKind::Lit(Literal::Int(f32_bits as i32 as i64)),
                span,
            )],
            span,
        );
    }
    make_call("f32_demote_f64", vec![v], span)
}

fn parse_float(s: &str) -> Expression {
    parse_float_at(s, false)
}

/// Parse a WAT float literal at the target width. `is_f32` rounds ONCE, to
/// single precision, and returns that value widened back to f64 — which is
/// exact, so the `f32_demote_f64` the lowering wraps it in changes nothing.
/// See `instr_float_is_f32` for why the width has to be known this early.
fn parse_float_at(s: &str, is_f32: bool) -> Expression {
    let s = s.trim();
    match s {
        "inf" | "+inf" => Expression::float(f64::INFINITY),
        "-inf" => Expression::float(f64::NEG_INFINITY),
        // All NaN forms — plain, `-nan`, `nan:0x…`, `nan:canonical`,
        // `nan:arithmetic`. Sign and payload are preserved; see `nan_literal`.
        _ if s.contains("nan") => Expression::float(nan_literal(s)),
        _ => {
            let cleaned = s.replace('_', "");
            if cleaned.contains("0x") || cleaned.contains("0X") {
                if let Some(v) = parse_hex_float_at(&cleaned, is_f32) {
                    return Expression::float(v);
                }
            }
            // Rust's own decimal parse is correctly rounded at each width, so
            // going through `f32` here is the decimal half of the same fix.
            Expression::float(if is_f32 {
                cleaned.parse::<f32>().unwrap_or(0.0) as f64
            } else {
                cleaned.parse::<f64>().unwrap_or(0.0)
            })
        }
    }
}

/// Parse a WAT hex float like `0x1.8p1` (= 1.5 × 2¹ = 3.0). Rust's `f64::parse`
/// rejects hex floats, so the significand (hex integer part + hex fraction) is
/// read here and scaled by the binary `p` exponent — but read EXACTLY, and
/// rounded once, for the reason spelled out inside.
fn parse_hex_float(s: &str) -> Option<f64> {
    parse_hex_float_at(s, false)
}

fn parse_hex_float_at(s: &str, is_f32: bool) -> Option<f64> {
    let (neg, rest) = match s.strip_prefix('-') {
        Some(r) => (true, r),
        None => (false, s.strip_prefix('+').unwrap_or(s)),
    };
    let rest = rest
        .strip_prefix("0x")
        .or_else(|| rest.strip_prefix("0X"))?;
    let (mantissa, exp) = match rest.find(['p', 'P']) {
        Some(i) => (&rest[..i], rest[i + 1..].parse::<i32>().ok()?),
        None => (rest, 0),
    };
    let (int_part, frac_part) = match mantissa.find('.') {
        Some(i) => (&mantissa[..i], &mantissa[i + 1..]),
        None => (mantissa, ""),
    };
    // The significand is read EXACTLY, into a 128-bit integer plus a sticky
    // flag, and rounded ONCE at the end.
    //
    // It used to be accumulated straight into an f64 — `value = value * 16 + d`
    // per digit — which rounds at every step. The spec has a whole "Rounding
    // behaviour" section precisely about literals whose significand is longer
    // than the format, and per-digit accumulation cannot answer it: by the time
    // the last digits arrive the result has already been rounded several times,
    // and the information that decided the answer is gone. `const.wast`'s
    // `+0x1.000000000000080000000000p-600`, which must round DOWN to
    // `0x1.0000000000000p-600`, and `+0x1.00000000000008000000000000000001p-600`,
    // which must round UP, differ only in bits the accumulation had already
    // discarded.
    //
    // 124 kept bits is far more than the 53 the format needs, so the single
    // rounding below sees everything that can affect it; anything past that
    // only has to be recorded as "nonzero", which is what `sticky` is.
    let mut mant: u128 = 0;
    let mut lsb_exp: i32 = 0; // the value is mant × 2^lsb_exp
    let mut sticky = false;
    let mut any_digit = false;
    // Digits arrive most-significant first, so once there is no room to shift,
    // every further digit is BELOW what we keep: it moves the kept bits' weight
    // up (integer part) or is simply dropped (fraction), and either way its
    // being nonzero is all that matters.
    for c in int_part.chars() {
        let d = c.to_digit(16)? as u128;
        any_digit = true;
        if mant <= (u128::MAX >> 4) {
            mant = (mant << 4) | d;
        } else {
            sticky |= d != 0;
            lsb_exp += 4;
        }
    }
    for c in frac_part.chars() {
        let d = c.to_digit(16)? as u128;
        any_digit = true;
        if mant <= (u128::MAX >> 4) {
            mant = (mant << 4) | d;
            lsb_exp -= 4;
        } else {
            sticky |= d != 0;
        }
    }
    if !any_digit {
        return None;
    }
    Some(round_binary(neg, mant, lsb_exp + exp, sticky, is_f32))
}

/// Round an EXACT binary value — `mant × 2^exp`, with `sticky` recording that
/// nonzero bits were dropped below `mant`'s last bit — to the nearest value of
/// the target format, ties to even. One rounding, at the end, which is what the
/// spec's rounding rules are stated in terms of. The f32 result comes back
/// WIDENED to f64, which is exact.
fn round_binary(neg: bool, mant: u128, exp: i32, sticky: bool, is_f32: bool) -> f64 {
    let signed_zero = if neg { -0.0f64 } else { 0.0f64 };
    if mant == 0 {
        return signed_zero;
    }
    // Significand bits (hidden bit included), the exponent of the smallest
    // subnormal, and the largest leading-bit exponent, per format.
    let (p, min_lsb, max_lead): (i32, i32, i32) = if is_f32 {
        (24, -149, 127)
    } else {
        (53, -1074, 1023)
    };
    let msb = 127 - mant.leading_zeros() as i32;
    // Below HALF the smallest subnormal nothing can round up to a nonzero
    // value. Returning here also keeps the shift below 129, so the masks are
    // always well defined.
    if exp + msb < min_lsb - 1 {
        return signed_zero;
    }
    // Keep `p` significand bits — but never place the last bit below the
    // subnormal grid, which is what makes gradual underflow come out right.
    let mut shift = msb + 1 - p;
    if exp + shift < min_lsb {
        shift = min_lsb - exp;
    }
    let (mut kept, mut kexp) = if shift <= 0 {
        // More bits available than we need: shifting LEFT is exact, and sticky
        // cannot be set here (it is only ever set once `mant` has filled up,
        // which forces a large positive shift).
        (mant << (-shift) as u32, exp + shift)
    } else {
        let s = shift as u32;
        let (dropped, keep, half) = if s >= 128 {
            // `half` = 2^(s-1) ≥ 2^127 and `mant` < 2^128, so only s == 128 can
            // even reach it; that is exactly a tie, which goes to the even
            // zero. Anything strictly above rounds up to the smallest value.
            (mant, 0u128, None)
        } else {
            (
                mant & ((1u128 << s) - 1),
                mant >> s,
                Some(1u128 << (s - 1)),
            )
        };
        let up = match half {
            Some(h) => dropped > h || (dropped == h && (sticky || (keep & 1) == 1)),
            None => s == 128 && mant > (1u128 << 127),
        };
        (keep + u128::from(up), exp + shift)
    };
    if kept == 0 {
        return signed_zero;
    }
    // Rounding up can carry out of the top bit (`0x1.fff…f` → `0x2.0`).
    if kept >> p != 0 {
        kept >>= 1;
        kexp += 1;
    }
    let lead = kexp + (127 - kept.leading_zeros() as i32);
    if lead > max_lead {
        return if neg {
            f64::NEG_INFINITY
        } else {
            f64::INFINITY
        };
    }
    // `kept` is at most 53 bits, so it converts to f64 exactly, and scaling by
    // a power of two is exact because the last bit is on the representable
    // grid — for the f32 case too, whose whole range sits inside f64's normals.
    // Step the exponent so no intermediate factor overflows: `2^1074` is
    // infinity, which is what used to make every subnormal literal —
    // `0x1p-1074`, the smallest f64 among them — parse as zero.
    let mut v = kept as f64;
    let mut e = kexp;
    while e > 0 {
        let step = e.min(1000);
        v *= 2f64.powi(step);
        e -= step;
    }
    while e < 0 {
        let step = e.max(-1000);
        v *= 2f64.powi(step);
        e -= step;
    }
    if neg { -v } else { v }
}

/// The BYTES a `(module binary "…" "…")` fixture spells.
///
/// ⚠ NOT `unquote`. That decodes a wast string as TEXT and knows only
/// `\n \t \r \\ \"`; a binary module is written almost entirely in HEX escapes
/// — `"\00asm" "\01\00\00\00"` — which `unquote` would pass through as the
/// four literal characters `\`, `0`, `0`, `a`. Feeding that to a decoder tests
/// nothing about the fixture.
///
/// The spec's string escapes for byte strings are `\XX` (two hex digits) plus
/// the same character escapes; anything else stands for its own UTF-8 bytes.
/// The pieces concatenate, because the spec splits long fixtures across several
/// literals.
fn binary_module_bytes(pair: &Pair<Rule>) -> Vec<u8> {
    let mut out: Vec<u8> = Vec::new();
    for s in pair.clone().into_inner() {
        if s.as_rule() != Rule::string {
            continue;
        }
        let raw = s.as_str().trim();
        let inner = raw
            .strip_prefix('"')
            .and_then(|r| r.strip_suffix('"'))
            .unwrap_or(raw);
        let bytes = inner.as_bytes();
        let mut i = 0usize;
        while i < bytes.len() {
            if bytes[i] != b'\\' {
                out.push(bytes[i]);
                i += 1;
                continue;
            }
            let Some(&next) = bytes.get(i + 1) else {
                out.push(b'\\');
                break;
            };
            match next {
                b'n' => (out.push(b'\n'), i += 2).1,
                b't' => (out.push(b'\t'), i += 2).1,
                b'r' => (out.push(b'\r'), i += 2).1,
                b'\\' => (out.push(b'\\'), i += 2).1,
                b'"' => (out.push(b'"'), i += 2).1,
                b'\'' => (out.push(b'\''), i += 2).1,
                _ => {
                    let hex = |b: u8| -> Option<u8> {
                        match b {
                            b'0'..=b'9' => Some(b - b'0'),
                            b'a'..=b'f' => Some(b - b'a' + 10),
                            b'A'..=b'F' => Some(b - b'A' + 10),
                            _ => None,
                        }
                    };
                    match (hex(next), bytes.get(i + 2).copied().and_then(hex)) {
                        (Some(hi), Some(lo)) => {
                            out.push((hi << 4) | lo);
                            i += 3;
                        }
                        // Not a valid escape — keep it verbatim rather than
                        // silently dropping a byte the fixture meant to include.
                        _ => {
                            out.push(b'\\');
                            i += 1;
                        }
                    }
                }
            }
        }
    }
    out
}

/// Decode a WAT string literal.
///
/// A single left-to-right scan, which the previous chain of `.replace()` calls
/// could not be. Two things were wrong with that chain:
///
///  * **No `\HH` hex escape.** It is the WAT spec's primary escape form
///    (`\00`–`\ff`), and dropping it left the two bytes in the text. That is
///    what made `comments.wast` unparseable: its `(module quote …)` ends line
///    comments with `\0a` / `\0d`, so with the escape undecoded the whole
///    module arrived as ONE line and every `;;` comment ran to the end of it.
///  * **Order dependence.** `.replace("\\n", "\n")` ran first and rewrote the
///    `\n` INSIDE a `\\n` (an escaped backslash followed by the letter n),
///    yielding backslash + newline instead of backslash + "n".
///
/// Bytes are accumulated and decoded at the end, since `\HH` can name a
/// non-ASCII byte; a lone invalid byte becomes U+FFFD rather than failing the
/// parse, matching how the rest of this front end treats malformed text.
fn unquote(s: &str) -> String {
    let s = s.trim();
    if !(s.len() >= 2 && s.starts_with('"') && s.ends_with('"')) {
        return s.to_string();
    }
    let inner = &s[1..s.len() - 1];
    let bytes = inner.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'\\' {
            out.push(bytes[i]);
            i += 1;
            continue;
        }
        // A trailing lone backslash: keep it rather than reading past the end.
        let Some(&c) = bytes.get(i + 1) else {
            out.push(b'\\');
            break;
        };
        match c {
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
            // `\u{...}` — a Unicode scalar, encoded as UTF-8.
            b'u' if bytes.get(i + 2) == Some(&b'{') => {
                match inner[i + 3..].find('}') {
                    Some(rel) => {
                        let hex = &inner[i + 3..i + 3 + rel];
                        match u32::from_str_radix(hex, 16).ok().and_then(char::from_u32) {
                            Some(ch) => {
                                let mut buf = [0u8; 4];
                                out.extend_from_slice(ch.encode_utf8(&mut buf).as_bytes());
                            }
                            // Not a scalar value: keep the text verbatim so a
                            // malformed-ness assertion still sees something wrong.
                            None => out.extend_from_slice(&bytes[i..i + 3 + rel + 1]),
                        }
                        i += 3 + rel + 1;
                    }
                    None => {
                        out.push(b'\\');
                        i += 1;
                    }
                }
            }
            // `\HH` — exactly two hex digits, one raw byte.
            _ => {
                let hi = (c as char).to_digit(16);
                let lo = bytes.get(i + 2).and_then(|&d| (d as char).to_digit(16));
                match (hi, lo) {
                    (Some(h), Some(l)) => {
                        out.push((h * 16 + l) as u8);
                        i += 3;
                    }
                    // Not a recognised escape: keep the backslash literally.
                    _ => {
                        out.push(b'\\');
                        i += 1;
                    }
                }
            }
        }
    }
    String::from_utf8_lossy(&out).into_owned()
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
            None => return Vec::new(),
        }
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

/// The tag a `throw`/`catch` names, in EITHER form. `peek_plain_label` bails on
/// a `folded_instr`, which is why the folded `(throw $t …)` had no tag at all.
fn peek_instr_tag_ref(__w: &mut WastWalker, pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    if !matches!(inner.as_rule(), Rule::plain_instr | Rule::folded_instr) {
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
        // An `instr_arg` spells a bare tagidx as `integer`, not `index` — the
        // `index` rule is only reached where the grammar names it explicitly
        // (a `try_clause`, say). Matching only `id`/`index` left `throw 0` with
        // no tag at all.
        .find(|c| matches!(c.as_rule(), Rule::id | Rule::index | Rule::integer))
        .map(|c| tag_ref_name(__w, c.as_str()))
}

/// The `(type $sig)` type name from a `call_indirect`/`return_call_indirect`
/// opener — its `block_type` immediate wraps a `(type index)` typeuse. Returns
/// the type index text (`$`-stripped) so its param count gives the call arity.
/// The `(argc, result count)` of a `call_indirect` / `return_call_indirect`
/// type use.
///
/// A type use is EITHER `(type $sig)` naming a declared function type OR the
/// inline `(param …)* (result …)*` spelling of the same shape — and the inline
/// form needs no `(type …)` at all. Only the named form was read, so
/// `call_indirect (result i32)` came out as 0→0 and the VM rejected a perfectly
/// good 0→1 callee with "indirect call type mismatch (callee 0→1, expected
/// 0→0)". When both are written the spec requires them to agree, so the named
/// one wins and the inline half is redundant.
fn peek_typeuse_shape(__w: &WastWalker, pair: &Pair<Rule>) -> (usize, usize) {
    let inner = match pair.as_rule() {
        Rule::instr => match pair.clone().into_inner().next() {
            Some(i) => i,
            None => return (0, 0),
        },
        _ => pair.clone(),
    };
    if !matches!(inner.as_rule(), Rule::plain_instr | Rule::folded_instr) {
        return (0, 0);
    }
    let mut params = 0usize;
    let mut results = 0usize;
    let mut named: Option<String> = None;
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
        // `block_type` inlines the keyword rather than nesting `param`/`result`
        // rules, so the discriminator is the head token of its own text.
        let head = bt.as_str().trim_start_matches('(').trim_start();
        let count = || {
            bt.clone()
                .into_inner()
                .filter(|x| matches!(x.as_rule(), Rule::val_type | Rule::any_val_type))
                .count()
        };
        if head.starts_with("type") {
            for x in bt.clone().into_inner() {
                if x.as_rule() == Rule::index {
                    named = Some(x.as_str().trim_start_matches('$').to_string());
                }
            }
        } else if head.starts_with("param") {
            params += count();
        } else if head.starts_with("result") {
            results += count();
        }
    }
    if let Some(n) = named {
        // ⚠ Type names are MODULE-QUALIFIED (`m#<seq>#<name>`) — a bare `$t`
        // would otherwise mean different types in different modules of one
        // script. A NUMERIC `(type 0)` is a declaration-order index instead, and
        // `type_index_name` already holds the qualified names.
        // Looking one up qualified and the other bare is not a near miss: it
        // reports `expected 0→1` for a 2→1 signature, which reads as a bad
        // table entry rather than a failed lookup.
        let key = match n.parse::<usize>() {
            Ok(i) => __w.type_index_name.get(i).cloned().unwrap_or(n),
            Err(_) => qualify_type_name(__w, &n),
        };
        return (
            __w.type_func_params.get(&key).copied().unwrap_or(0) as usize,
            __w.type_func_results.get(&key).copied().unwrap_or(0) as usize,
        );
    }
    (params, results)
}

fn peek_typeuse_index(pair: &Pair<Rule>) -> Option<String> {
    let inner = if pair.as_rule() == Rule::instr {
        pair.clone().into_inner().next()?
    } else {
        pair.clone()
    };
    // The `(type $sig)` type-use nests identically (`instr_arg → block_type`)
    // in the plain and folded forms.
    if !matches!(inner.as_rule(), Rule::plain_instr | Rule::folded_instr) {
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
    if !matches!(inner.as_rule(), Rule::plain_instr | Rule::folded_instr) {
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

fn fold_instructions(__w: &mut WastWalker, 
    pairs: Vec<Pair<Rule>>,
    labels: &mut LabelStack,
) -> Result<Vec<Statement>, String> {
    fold_instructions_seeded(__w, pairs, labels, Vec::new())
}

/// Like `fold_instructions`, but the value stack starts pre-loaded with `seed`
/// (bottom-to-top). Used to thread WASM `block (param …)` inputs into the body.
fn fold_instructions_seeded(__w: &mut WastWalker,
    // `mut` because a PLAIN instruction that absorbed a following folded one
    // splices it back in as the next instruction — see the `plain_instr` arm.
    mut pairs: Vec<Pair<Rule>>,
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
                        let result_temps: Vec<String> = (0..peek_block_result_count(__w, &pairs[i]))
                            .map(|_| fresh_result_temp(__w))
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
                        preserve_stack_across_block(__w, &mut stack, &mut statements);
                        let then_end = else_idx.unwrap_or(end_idx);
                        // Anything the opener swallowed belongs to the THEN
                        // branch, ahead of what follows the opener token.
                        let mut then_pairs: Vec<Pair<Rule>> = opener_folded_instrs(&pairs[i]);
                        then_pairs.extend(pairs[i + 1..then_end].iter().cloned());
                        labels.push(__w, label.clone(), LabelKind::Block, Vec::new());
                        let mut then_body =
                            fold_instructions_seeded(__w, then_pairs, labels, seed.clone())?;
                        let mut else_body = if let Some(ei) = else_idx {
                            // `else` is a `plain_instr` too, so it swallows the
                            // first folded instruction of ITS branch the same
                            // way the opener does.
                            let mut else_pairs: Vec<Pair<Rule>> = opener_folded_instrs(&pairs[ei]);
                            else_pairs.extend(pairs[ei + 1..end_idx].iter().cloned());
                            Some(fold_instructions_seeded(__w, else_pairs, labels, seed)?)
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
                                        with_events: false,
                                    }],
                                    kind: VarDeclKind::Let,
                                }));
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
                                    elifs: Vec::new(),
                                },
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
                        let seed_vals = if param_count > 0 && stack.len() >= param_count {
                            stack.split_off(stack.len() - param_count)
                        } else {
                            Vec::new()
                        };
                        preserve_stack_across_block(__w, &mut stack, &mut statements);
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
                            (0..param_count).map(|_| fresh_result_temp(__w)).collect()
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
                                        with_events: false,
                                    }],
                                    kind: VarDeclKind::Let,
                                }));
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
                        let result_temps: Vec<String> = (0..peek_block_result_count(__w, &pairs[i]))
                            .map(|_| fresh_result_temp(__w))
                            .collect();
                        for tmp in &result_temps {
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
                        let effective = labels.push(__w, label.clone(), kind, result_temps.clone());
                        labels.set_last_param_temps(param_temps.clone());
                        let mut body = fold_instructions_seeded(__w, body_pairs, labels, seed)?;
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
                    let tag = peek_instr_tag_ref(__w, &pairs[i]).unwrap_or_default();
                    let arity = tag_arity(__w, &tag) as usize;
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
                            let tmp = fresh_result_temp(__w);
                            statements.push(Statement::new(StmtKind::VarDecl {
                                declarations: vec![VarDeclarator {
                                    pattern: BindingPattern::Ident(tmp.clone()),
                                    type_hint: None,
                                    init: Some(exnref_expr),
                                    array_bounds: None,
                                    with_events: false,
                                }],
                                kind: VarDeclKind::Let,
                            }));
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
                    emit_folded_br_on_null(__w, 
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
                    // argc and the expected result count — the two halves of
                    // the type shape the VM checks the funcref against (traps
                    // on mismatch) — from `(type $sig)` or the inline
                    // `(param …)(result …)` spelling.
                    let (argc, expected_results) = peek_typeuse_shape(__w, &pairs[i]);
                    // Optional table reference `call_indirect $t (type $sig)`
                    // dispatches through a NAMED table (default table 0).
                    let tableidx = peek_call_indirect_table(&pairs[i])
                        .map(|t| resolve_table_index(__w, &t) as usize)
                        .unwrap_or_else(|| __w.table_index_base);
                    let n = (argc + 1).min(stack.len());
                    let operands: Vec<Expression> = stack.split_off(stack.len() - n);
                    // ⛔ THE DECLARED FUNCTYPE RIDES ALONG HERE TOO. The folded
                    // arm appended it and this one did not, so every PLAIN
                    // `call_indirect (type $t)` reached the VM with the three
                    // counts and no type — and the runtime check, which needs
                    // the type to tell `(func (result i32))` from
                    // `(func (result i64))`, had nothing to compare and fell
                    // back to arity. One instruction, two lowerings, one of
                    // them fixed.
                    let signature = typeuse_signature(__w, &pairs[i]);
                    let canon = typeuse_canon_name(__w, &pairs[i]);
                    let mut call_args = vec![
                        Expression::int(argc as i64),
                        Expression::int(tableidx as i64),
                        Expression::int(expected_results as i64),
                        Expression::string(&signature),
                        Expression::string(&canon),
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
                        // See the folded arm: the destructure belongs with a
                        // compiler change that is not made, so the packed push
                        // stays.
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
                        emit_folded_block(__w, inner, false, span, labels, &mut statements, &mut stack)?;
                    }
                    "loop" => {
                        emit_folded_block(__w, inner, true, span, labels, &mut statements, &mut stack)?;
                    }
                    "if" => {
                        emit_folded_if(__w, inner, span, labels, &mut statements, &mut stack)?;
                    }
                    "try_table" => {
                        emit_folded_try_table(__w, inner, span, labels, &mut statements, &mut stack)?;
                    }
                    "br_on_null" => {
                        emit_folded_br_on_null(__w, 
                            inner,
                            false,
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        )?;
                    }
                    "br_on_non_null" => {
                        emit_folded_br_on_null(__w, 
                            inner,
                            true,
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        )?;
                    }
                    "return" => {
                        emit_folded_return(__w, inner, span, labels, &mut statements, &mut stack)?;
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
                                emit_folded_br_on_null(__w, 
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
                                emit_folded_br_on_null(__w, 
                                    op,
                                    true,
                                    span,
                                    labels,
                                    &mut statements,
                                    &mut stack,
                                )?;
                                stack.pop();
                            }
                            // A dropped operand that branches or is structured
                            // (`(drop (br_if $L (block …) (cond)))`) must run
                            // through the statement machinery; the expression
                            // walk would discard the branch. Evaluate it onto
                            // the stack, then pop for the drop — landing the
                            // popped value as a statement, since it may be a
                            // DEFERRED expression (e.g. a call) whose side
                            // effects must still run in program order.
                            (Some(op), _) if folded_needs_stmt_lowering(&op) => {
                                emit_folded_operand_stmtwise(__w, 
                                    op,
                                    labels,
                                    &mut statements,
                                    &mut stack,
                                )?;
                                if let Some(v) = stack.pop() {
                                    statements.push(Statement::with_span(StmtKind::Expr(v), span));
                                }
                            }
                            // A BARE `(drop)` — no folded operand — takes its
                            // value from the enclosing STACK, exactly like the
                            // plain `drop` form (folding is sugar). Routing it
                            // through the expression walk instead made
                            // `walk_folded_core` substitute null
                            // (`args.into_iter().next().unwrap_or(null)`), so the
                            // drop consumed a NULL and the real value survived:
                            // `(block (result i32 i32 i32) …) (drop) (drop)` left
                            // the LAST block result instead of the first
                            // (block/loop/if.wast `"multi"`; traced as
                            // `ref.null; drop` twice). The popped value is landed
                            // as a statement because it may be a DEFERRED
                            // expression whose side effects must still run.
                            // This mirrors the `has_nested_operand` rule below,
                            // which drop never reaches — it is claimed here first.
                            (None, _) if !stack.is_empty() => {
                                let v = stack.pop().expect("non-empty");
                                statements.push(Statement::with_span(StmtKind::Expr(v), span));
                            }
                            _ => {
                                // Ordinary drop: a void instr, emit in program order.
                                let expr = walk_folded_instr_as_expr(__w, inner, span, labels)?;
                                statements.push(Statement::with_span(StmtKind::Expr(expr), span));
                            }
                        }
                    }
                    _ => {
                        // A branching folded instr (`br_if` at the head, or a
                        // `br`/`br_if`/`br_table`/`return` nested in an operand,
                        // e.g. `(i32.ctz (br_if 0 (i32.const 1) (i32.const 1)))`)
                        // — or one with a structured operand (`(i32.mul … (block
                        // …))`) — cannot go through the expression walk: the
                        // branch/label machinery would lower to null and vanish.
                        // Unfold it through the statement machinery instead.
                        if folded_needs_stmt_lowering(&inner) {
                            emit_folded_stmtwise(__w, inner, span, labels, &mut statements, &mut stack)?;
                            continue;
                        }
                        // A folded operand nests as `instr` or `instr_arg →
                        // folded/plain_instr`. When a stack-consuming instr has
                        // NONE of its value operands nested (the abbreviated
                        // "flat" style, e.g. `(struct.get $t 0)` after a block
                        // that left the ref on the stack), it takes its operands
                        // from the enclosing stack exactly like the plain form.
                        // Without this the missing operand became null.
                        // COUNT them rather than only asking whether any exist.
                        // The abbreviation is not all-or-nothing: an instruction
                        // may nest FEWER operands than it consumes, and the
                        // remainder come from the enclosing stack, BELOW the
                        // nested ones. Only the zero-nested case was handled
                        // here, so `(i32.const 7) (i32.add (i32.const 100))` —
                        // one nested operand against an arity of 2 — dropped the
                        // 7 and answered 100 instead of 107. `push_folded_operand`
                        // has always done this correctly for an operand POSITION;
                        // the sequence walker did not.
                        let nested_count = inner
                            .clone()
                            .into_inner()
                            .filter(|c| folded_operand_child(c).is_some())
                            .count();
                        // ⛔ A NESTED OPERAND IS NOT AN IMMEDIATE. `instr_arg`
                        // wraps both, so without the `folded_operand_child`
                        // exclusion — the one `push_folded_operand` has always
                        // had — every nested operand was walked TWICE: once as
                        // an argument expression here, and again by the loop
                        // below that pushes it on the stack.
                        //
                        // A binary op survived that by accident (its extra
                        // trailing arg is ignored and the first two happen to
                        // be right). `call` does not: the duplicate shifts every
                        // operand, so `(call $swap (call $swap …))` passed the
                        // inner CALL as argument 1 and the spread temps after
                        // it — the callee got three arguments and the inner
                        // call was evaluated twice.
                        let immediate_args: Vec<Expression> = inner
                            .clone()
                            .into_inner()
                            .filter(|c| {
                                c.as_rule() == Rule::instr_arg && folded_operand_child(c).is_none()
                            })
                            .map(|c| walk_instr_arg_pair(__w, c, labels))
                            .collect::<Result<_, _>>()?;
                        let arity = get_instruction_arity(__w, &head, &immediate_args);
                        if nested_count < arity && (arity > 0 || head == "call") {
                            let mut args = immediate_args;
                            // Take the shortfall from the stack FIRST — those
                            // are the operands the flat form would have found
                            // already there, and they sit below the nested ones.
                            let pop_count = (arity - nested_count).min(stack.len());
                            let drain_start = stack.len() - pop_count;
                            let popped: Vec<Expression> = stack.drain(drain_start..).collect();
                            args.extend(popped);
                            // Then evaluate the nested operands, in source
                            // order, on top of them.
                            let base = stack.len();
                            for child in inner.clone().into_inner() {
                                if let Some(nested) = folded_operand_child(&child) {
                                    if nested.as_rule() == Rule::folded_instr {
                                        push_folded_operand(
                                            __w,
                                            nested,
                                            labels,
                                            &mut statements,
                                            &mut stack,
                                        )?;
                                    } else {
                                        let s = to_span(&nested);
                                        stack.push(walk_plain_instr_as_expr(__w, nested, s, labels)?);
                                    }
                                }
                            }
                            args.extend(stack.drain(base..).collect::<Vec<_>>());
                            // A `call` yields as many values as the callee has
                            // results — a void call is a statement in program
                            // order, exactly like the plain form.
                            let pushes = if head == "call" {
                                call_result_count(__w, &args)
                            } else {
                                get_instruction_push_count(&head)
                            };
                            let expr = map_instr_to_ast(__w, head.clone(), args, span)?;
                            land_instr_value(__w, 
                                expr,
                                pushes,
                                head == "call",
                                span,
                                &mut statements,
                                &mut stack,
                            );
                            continue;
                        }
                        // A fully-folded instr is self-contained (all operands
                        // nested), so statement-vs-stack is purely an ordering
                        // question. Void instructions (`local.set`, stores,
                        // `struct.set`/`array.set`, bulk ops) must run in program
                        // order — deferring them on the value stack lets a later
                        // reader observe the pre-write state. Value producers stay
                        // on the stack for their consumer; a `call` lands by its
                        // callee's declared result count (0 = void statement).
                        let pushes = if head == "call" {
                            call_result_count(__w, &immediate_args)
                        } else {
                            get_instruction_push_count(&head)
                        };
                        let expr = walk_folded_instr_as_expr(__w, inner, span, labels)?;
                        if head.is_empty() {
                            stack.push(expr);
                        } else {
                            land_instr_value(__w, 
                                expr,
                                pushes,
                                head == "call",
                                span,
                                &mut statements,
                                &mut stack,
                            );
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
                let name = peel_mem_selector(__w, &name, &mut raw_args, labels)?;

                // A PLAIN instruction takes no folded operands: folding happens
                // only inside an explicit `( … )`, so `drop (i32.const 5)` is
                // `drop` followed by `i32.const 5`, two instructions. The
                // grammar's `instr_arg*` absorbs the following folded
                // instruction into the plain one anyway, so split it back out
                // here and hand it to the sequence AFTER this instruction runs.
                //
                // `nop` had this fixed individually; the same absorption
                // affects every plain instruction, and `drop`/`local.set` were
                // getting the absorbed value as an OPERAND — `7 8 drop
                // (i32.const 5) i32.add` answered 8 instead of 12.
                //
                // The reftype immediate `(ref null? ht)` is NOT an absorbed
                // instruction; it has its own `ref_type_arg` rule (ahead of
                // `folded_instr` in `instr_arg`) precisely so it stays here.
                let mut absorbed: Vec<Pair<Rule>> = Vec::new();
                let mut immediates: Vec<Pair<Rule>> = Vec::new();
                for raw in raw_args {
                    match raw
                        .clone()
                        .into_inner()
                        .next()
                        .filter(|x| x.as_rule() == Rule::folded_instr)
                    {
                        // Keep the INNER `folded_instr`: the sequence loop uses
                        // a non-`instr` pair directly as its `inner`, so the
                        // `instr_arg` wrapper would not match the folded arm.
                        Some(folded) => absorbed.push(folded),
                        None => immediates.push(raw),
                    }
                }
                // Splice them in as the NEXT instructions. `i` was already
                // advanced past the current pair, so index `i` is the slot
                // immediately after it, and this survives every `continue` in
                // the arms below.
                for (k, p) in absorbed.into_iter().enumerate() {
                    pairs.insert(i + k, p);
                }

                // Parse inline arguments
                let mut args = Vec::new();
                for raw in immediates {
                    args.push(walk_instr_arg_for(__w, raw, labels, &name)?);
                }

                // Determine stack arity
                let arity = get_instruction_arity(__w, &name, &args);
                let pop_count = usize::min(arity, stack.len());
                let drain_start = stack.len() - pop_count;
                let popped: Vec<Expression> = stack.drain(drain_start..).collect();

                // Append popped operands to args
                args.extend(popped);

                // Handle control flow instructions that are statements
                match name.as_str() {
                    "nop" => {
                        // `nop` does nothing — but it is not free to ignore its
                        // `args`. The spec's text format has no operands on a
                        // PLAIN instruction: a folded instruction written after
                        // one is the NEXT instruction in the sequence
                        // (`nop (i32.const 5)` is `nop` then `i32.const 5`).
                        // `instr_arg` absorbs it into the `nop` anyway, so
                        // discarding `args` DELETED it — the block came out
                        // empty and produced null. `nop` pops nothing, so
                        // everything here is an absorbed instruction: hand the
                        // values back to the sequence, in order.
                        for value in args {
                            stack.push(value);
                        }
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
                        statements.push(Statement::with_span(StmtKind::Expr(trap_expr()), span));
                        stack.push(trap_expr());
                    }
                    "return" => {
                        let n = __w.current_fn_results;
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
                        // See the folded arm: the ref form has its own opcode
                        // and a funcref VALUE, not a callee NAME.
                        let inner = if name == "return_call" {
                            "call"
                        } else {
                            "return_call_ref"
                        };
                        let call = map_instr_to_ast(__w, inner.to_string(), args, span)?;
                        if inner == "return_call_ref" {
                            statements.push(Statement::with_span(StmtKind::Expr(call), span));
                        } else if let ExprKind::Call {
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
                        emit_br_stmt_carry(__w, args.first(), span, labels, &mut statements, &mut stack);
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
                        emit_br_if_stmt(__w, 
                            lbl_arg,
                            cond_expr,
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        );
                    }
                    "br_table" => {
                        emit_br_table_stmt(__w, &args, span, labels, &mut statements, &mut stack);
                    }
                    // `br_on_cast L $from $to` branches to L (carrying the ref as
                    // the block result) when the ref IS `$to`; `br_on_cast_fail`
                    // when it is NOT. The ref stays on the stack for the
                    // fall-through path (like `br_if`'s peeked block result).
                    "br_on_cast" | "br_on_cast_fail" => {
                        emit_br_on_cast_stmt(__w,
                            &args,
                            name == "br_on_cast_fail",
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        );
                    }
                    "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail" => {
                        emit_br_on_cast_desc_eq_stmt(
                            __w,
                            &args,
                            name == "br_on_cast_desc_eq_fail",
                            span,
                            labels,
                            &mut statements,
                            &mut stack,
                        );
                    }
                    _ => {
                        // A `call` yields as many values as the callee has results;
                        // a 0-result (void) call is a statement that must run in
                        // place, not a deferred stack value. Everything else that
                        // reaches here pushes a single value.
                        let pushes = if name == "call" {
                            call_result_count(__w, &args)
                        } else {
                            get_instruction_push_count(&name)
                        };
                        let expr = map_instr_to_ast(__w, name.clone(), args, span)?;
                        // ONE landing rule for flat and folded alike (this arm
                        // used to inline its own copy, which then missed the
                        // void-ordering rule the folded path had — stack.wast
                        // `not-quite-a-tree` returned 5 instead of 3).
                        land_instr_value(__w, 
                            expr,
                            pushes,
                            name == "call",
                            span,
                            &mut statements,
                            &mut stack,
                        );
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

fn get_instruction_arity(__w: &mut WastWalker, name: &str, args: &[Expression]) -> usize {
    // The `@@off<N>` / `@@mem<N>` suffixes are emitter channels, not part
    // of the op identity: the base name is everything before the first `@@`.
    let name = name.split_once("@@").map(|(b, _)| b).unwrap_or(name);
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

        // `nop` pops nothing. It was falling to the `_ => 1` default and
        // stealing a value off the operand stack, which the `"nop"` arm then
        // discarded — `get_instruction_push_count` already listed it at 0, so
        // the two arities disagreed.
        "nop" => 0,

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
        // The descriptor-comparing casts pop TWO: the reference and the
        // descriptor to compare it against. The reftype stays an immediate.
        "ref.cast_desc_eq" => 2,
        // Same two stack operands as `ref.cast_desc_eq` — the reference and
        // the descriptor. The label and BOTH reftypes are immediates.
        "br_on_cast_desc_eq" | "br_on_cast_desc_eq_fail" => 2,

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
                        *__w.func_name_arities.get(n).unwrap_or(&1)
                    }
                    ExprKind::Lit(Literal::Int(idx)) => {
                        *__w.func_index_arities.get(*idx as usize).unwrap_or(&1)
                    }
                    _ => 1,
                }
            } else {
                1
            }
        }

        // A tagged call pops the funcref plus the TAG's params — the tag
        // carries the signature here, not a `$sig` immediate.
        "call_with_tag" | "call_return_with_tag" => {
            let params = args
                .first()
                .map(wasm_type_ref_name)
                .and_then(|n| __w.call_tag_params.get(&n).map(|(p, _)| *p))
                .unwrap_or(0);
            1 + params
        }
        // `$table $call_tag` are both immediates; the element index is the one
        // stack operand, plus the tag's params.
        "call_indirect_with_tag" => {
            let params = args
                .get(1)
                .map(wasm_type_ref_name)
                .and_then(|n| __w.call_tag_params.get(&n).map(|(p, _)| *p))
                .unwrap_or(0);
            1 + params
        }

        "call_indirect" | "return_call_indirect" => 2,
        // call_ref pops the funcref plus the sig's params.
        "call_ref" | "return_call_ref" => {
            // Both spellings of `$sig` — a numeric type index reaches here too.
            let t = resolve_wast_type_name(__w, args.first());
            1 + __w.type_func_params.get(&t).copied().unwrap_or(0)
        }

        // GC struct ops
        // struct.new $T: typeidx is an immediate; field values come from stack.
        // We stored field counts by type name in STRUCT_FIELD_COUNTS.
        "struct.new" => {
            // args[0] is typeidx immediate (ident or int) — not a stack value.
            // Remaining stack operands = field count for that type.
            if let Some(first) = args.first() {
                let _ = first;
                let type_name = resolve_wast_type_name(__w, args.first());
                *__w.struct_field_counts.get(&type_name).unwrap_or(&0)
            } else {
                0
            }
        }
        "struct.new_default" => 0, // no stack operands; typeidx is immediate
        // Custom Descriptors: the descriptor is a STACK operand, pushed on top
        // of the field values, so these are `struct.new`'s arity plus one.
        // Without an entry here the fold logic reordered the operands around
        // the instruction and the type immediate was never resolved.
        "struct.new_desc" => {
            if let Some(first) = args.first() {
                let _ = first;
                let type_name = resolve_wast_type_name(__w, args.first());
                *__w.struct_field_counts.get(&type_name).unwrap_or(&0) + 1
            } else {
                1
            }
        }
        "struct.new_default_desc" => 1, // descriptor only
        "ref.get_desc" => 1,            // [ref] → [descriptor]
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
fn call_result_count(__w: &mut WastWalker, args: &[Expression]) -> usize {
    match args.first().map(|e| &e.kind) {
        Some(ExprKind::Ident(n)) => __w.func_name_results.get(n).copied()
            .unwrap_or(1),
        Some(ExprKind::Lit(Literal::Int(idx))) => __w.func_index_results.get(*idx as usize).copied()
            .unwrap_or(1),
        _ => 1,
    }
}

fn get_instruction_push_count(name: &str) -> usize {
    // The `@@off<N>` / `@@mem<N>` suffixes are emitter channels, not part
    // of the op identity: the base name is everything before the first `@@`.
    let name = name.split_once("@@").map(|(b, _)| b).unwrap_or(name);
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

// ═══ THE STACK-TYPING PASS ═══════════════════════════════════════════════════
//
// WASM §3.3 validation over the MODULE TEXT: an abstract operand stack and a
// control stack, run per function body. This is the general pass that
// `descriptor_operand_mismatch` above deliberately refused to grow into — it
// is the 2297 "type mismatch" assertions, 84% of the suite's `assert_invalid`
// corpus, and it is one algorithm rather than a rule per fixture.
//
// ⛔ THE BAIL IS THE SAFETY PROPERTY, INHERITED VERBATIM. Anything this pass
// cannot type — an unmodelled mnemonic, an unresolvable index, a construct
// outside its scope — abandons the whole function and reports NOTHING, so the
// assertion stays honestly red. `assert_invalid` compares the expected
// diagnostic, so a pass that guesses turns assertions green for reasons it
// never established. Unknown must propagate, never default.
//
// ⛔ IT RUNS LAST. Presence before agreement: a module with `local.get 99` AND
// a type error must report "unknown local". The name-resolution, const-expr,
// limits and immutable-global rules all settle before this one is asked.

/// A function type, resolved to the typing lattice.
#[derive(Clone, Debug)]
struct FnSig {
    params: Vec<Vt>,
    results: Vec<Vt>,
}

/// Why a function body stopped being typed.
enum Fail {
    /// A rule the module provably violates — reportable.
    Mismatch(String),
    /// Not typeable by this pass. Reports nothing.
    Bail,
}
type R<T> = std::result::Result<T, Fail>;

fn tmis<T>() -> R<T> {
    Err(Fail::Mismatch("type mismatch".to_string()))
}

/// "unknown global 1" for a numeric spelling, bare "unknown global" for a
/// `$name`. The suite asserts the index where the source writes one, and the
/// comparison is one-directional — a reason vaguer than the fixture does not
/// discharge it.
fn unknown_index_msg(kind: &str, key: &str) -> String {
    match key.trim().parse::<u64>() {
        Ok(n) => format!("unknown {kind} {n}"),
        Err(_) => format!("unknown {kind}"),
    }
}

/// The same failure, carrying WHAT disagreed.
///
/// ⛔ THIS IS NOT COSMETIC AND IT DOES NOT WEAKEN THE ASSERTION. The suite's
/// comparison is one-directional — our reason must CONTAIN the asserted text —
/// so "type mismatch: …" discharges a `"type mismatch"` fixture exactly as the
/// bare string does. What it buys is the only thing that can tell an
/// over-firing rule from a correct one: 2297 fixtures all assert the same two
/// words, so a pass that fires for the WRONG reason is indistinguishable from
/// one that fires for the right one unless the message says which.
fn tmis_d<T>(detail: String) -> R<T> {
    Err(Fail::Mismatch(format!("type mismatch: {detail}")))
}

/// A value type in its spec spelling.
fn vt_show(v: &Vt) -> String {
    match v {
        Vt::Num(n) => (*n).to_string(),
        Vt::Bottom => "bot".to_string(),
        Vt::Ref(r) => {
            let h = match &r.heap {
                Heap::Abs(a) => (*a).to_string(),
                Heap::Concrete(i) => format!("{i}"),
            };
            format!(
                "(ref{}{})",
                if r.nullable { " null" } else { "" },
                if r.exact { format!(" (exact {h})") } else { format!(" {h}") }
            )
        }
    }
}

/// Which types may serve as the target of the IMPLICIT-type dedup.
///
/// §6.6.4 lets an inline signature REUSE an existing type instead of defining a
/// new one — but reuse means *the same type*, and type identity is
/// ISO-RECURSIVE. A type declared inside a multi-member `(rec …)` is identified
/// by its whole group, so a standalone inline signature is never the same type
/// as one of its members however well the parameters and results line up.
///
/// ⛔ THIS IS WHAT `type-rec.wast` ASSERTS, AND A SIGNATURE MATCH CANNOT SEE IT.
/// `(rec (type $ft (func)) (type (func)))` with a plain `(func $f)` gave `$f`
/// the index of `$ft` by structural match, so `(global (ref $ft) (ref.func $f))`
/// compared a type against itself and typed clean. The three fixtures differ
/// from their valid neighbours in the REC GROUP alone.
///
/// A standalone `(type …)` is its own singleton group, so the common case is
/// unaffected.
fn implicit_dedup_eligible(types: &[DescType]) -> Vec<bool> {
    let mut size: HashMap<usize, usize> = HashMap::new();
    for t in types {
        *size.entry(t.rec_group).or_insert(0) += 1;
    }
    types
        .iter()
        .map(|t| size.get(&t.rec_group).copied().unwrap_or(1) == 1)
        .collect()
}

/// Everything a body is typed against: the module's index spaces.
#[derive(Default)]
struct TypeCtx {
    types: Vec<DescType>,
    type_names: HashMap<String, usize>,
    /// Per type index: its function signature, when it is a functype.
    type_sigs: Vec<Option<FnSig>>,
    funcs: Vec<Option<FnSig>>,
    func_names: HashMap<String, usize>,
    /// Per function index: the type index it was declared with, for `ref.func`.
    func_types: Vec<Option<usize>>,
    /// ⛔ `Option`, NOT `Vt::Bottom`. Bottom is the POLYMORPHIC bottom — a
    /// subtype of everything, produced by `unreachable` — and `pop_expect`
    /// short-circuits on it in BOTH directions. Reusing it for "this spelling
    /// did not parse" gave one sentinel two meanings, so an unparseable
    /// declared type silently accepted every value instead of abstaining.
    /// `None` is the honest answer and its consumers bail on it.
    globals: Vec<Option<Vt>>,
    /// Per global index: whether it was declared `(mut t)`.
    ///
    /// ⛔ THIS RIDES THE INDEX SPACE ON PURPOSE. The rule already existed as a
    /// module-level scan keyed on the global's NAME, which meant
    /// `global.set $g` was checked and `global.set 0` was not, and an IMPORTED
    /// global — which occupies the same index space but is not a
    /// `global_field` — was never collected at all. Both spellings resolve
    /// through `global_names`/`globals` here, so both are checked by
    /// construction.
    global_mut: Vec<bool>,
    /// How many globals are IMPORTED — the prefix of the index space a
    /// constant expression is allowed to read. See `Tv::global_limit`.
    imported_globals: usize,
    global_names: HashMap<String, usize>,
    /// Per memory: its address type. memory64 declares `i64`; the default is
    /// `i32`, and it decides the operand type of every load, store and
    /// `memory.*` on that memory.
    mem_addr: Vec<&'static str>,
    tables: Vec<Option<Vt>>,
    table_addr: Vec<&'static str>,
    table_names: HashMap<String, usize>,
    tags: Vec<Vec<Vt>>,
    tag_names: HashMap<String, usize>,
    /// Per element segment: its declared reference type.
    ///
    /// ⛔ A SEGMENT IS NOT ITS TABLE. `array.init_elem` compares the SEGMENT's
    /// type against the array's element type, and an active segment's table
    /// says nothing about a passive one — which is the only kind the fixtures
    /// use. `None` where the spelling did not parse, so the check abstains
    /// rather than guessing.
    elem_types: Vec<Option<Vt>>,
    elem_names: HashMap<String, usize>,
}

/// One flattened instruction. Folded and plain source forms both reduce to
/// this sequence — which is the point: the typing rules are written once.
struct FlatInstr<'a> {
    head: String,
    imms: Vec<String>,
    /// The `block_type` pairs when this instruction opens a control frame,
    /// and the `(result …)` of a typed `select`.
    btypes: Vec<Pair<'a, Rule>>,
    label: Option<String>,
    /// `try_table`'s `try_clause*`. Each names a branch target, and the
    /// `catch`/`catch_ref` forms name a tag as well — the values the branch
    /// carries come from BOTH, so the clause has to travel with the frame.
    clauses: Vec<Pair<'a, Rule>>,
}

/// A control frame, per the spec's validation algorithm.
struct Ctrl {
    op: &'static str,
    label: Option<String>,
    start: Vec<Vt>,
    end: Vec<Vt>,
    height: usize,
    unreachable: bool,
    /// The local-init state when this frame OPENED.
    ///
    /// ⛔ INITS MADE INSIDE A STRUCTURED INSTRUCTION DO NOT ESCAPE IT, and
    /// `local_init.wast` pins that harder than it looks: setting the local in
    /// BOTH the `then` and the `else` and reading it after the `if` is still
    /// "uninitialized local". A join over the branches — the intuitive reading
    /// — accepts that module and fails the fixture.
    inited_entry: Vec<bool>,
}

struct Tv<'a> {
    ctx: &'a TypeCtx,
    vals: Vec<Vt>,
    ctrls: Vec<Ctrl>,
    locals: Vec<Vt>,
    local_names: HashMap<String, usize>,
    /// Per local: has it been assigned on every path to here?
    inited: Vec<bool>,
    /// How much of the global index space this expression may read.
    ///
    /// §3.4.10 types a module's constant expressions in a context whose
    /// globals are the IMPORTED ones only, so `(global $a i32 (global.get $b))`
    /// is "unknown global" even when `$b` is declared — before OR after it.
    /// The suite pins both halves: every valid init-expr `global.get` in the
    /// corpus names an import, and a table initializer reading a defined
    /// global is asserted invalid.
    ///
    /// `None` in a function body, where the whole space is visible.
    global_limit: Option<usize>,
}

impl<'a> Tv<'a> {
    /// The constant instructions, per §3.4.10 plus the extended-const and GC
    /// additions that WASM 3.0 folds in.
    ///
    /// ⛔ `i32.add`/`sub`/`mul` ARE CONSTANT NOW, and the suite relies on it:
    /// `elem.wast:1062` and `data.wast:180` are VALID modules whose offsets
    /// are `(i32.add (i32.const 1) (i32.const 2))`. Writing the pre-3.0 set
    /// here would have rejected them — the direction that costs working
    /// modules rather than a missed diagnostic. `i32.ctz`, which the suite
    /// asserts invalid in the same position, is the line: arithmetic that
    /// extended-const names, nothing else.
    fn is_const_instr(name: &str) -> bool {
        name.ends_with(".const")
            || matches!(
                name,
                "end"
                    | "ref.null"
                    | "ref.func"
                    | "global.get"
                    | "ref.i31"
                    | "struct.new"
                    | "struct.new_default"
                    | "array.new"
                    | "array.new_default"
                    | "array.new_fixed"
                    | "any.convert_extern"
                    | "extern.convert_any"
                    | "i32.add"
                    | "i32.sub"
                    | "i32.mul"
                    | "i64.add"
                    | "i64.sub"
                    | "i64.mul"
            )
    }

    // ── the spec's algorithm, verbatim ───────────────────────────────────
    //
    // ⛔ `pop_val` is the detail that decides `unreached-invalid.wast` (117
    // assertions). After `unreachable`, pops return the BOTTOM type instead of
    // erroring — but the enclosing frame's height floor is still enforced.
    // Treat unreachable code as unvalidated and all 117 report "the module
    // validated"; treat it as unconstrained and they fail the other way.

    fn pop_val(&mut self) -> R<Vt> {
        let f = self.ctrls.last().ok_or(Fail::Bail)?;
        if self.vals.len() == f.height {
            if f.unreachable {
                return Ok(Vt::Bottom);
            }
            return tmis_d("stack underflow in block".to_string());
        }
        self.vals.pop().ok_or(Fail::Bail)
    }

    fn pop_expect(&mut self, e: &Vt) -> R<Vt> {
        let a = self.pop_val()?;
        if matches!(a, Vt::Bottom) {
            return Ok(e.clone());
        }
        if matches!(e, Vt::Bottom) {
            return Ok(a);
        }
        if !vt_subtype(&a, e, &self.ctx.types) {
            return tmis_d(format!("expected {}, got {}", vt_show(e), vt_show(&a)));
        }
        Ok(a)
    }

    fn push_val(&mut self, v: Vt) {
        self.vals.push(v);
    }

    fn push_vals(&mut self, vs: &[Vt]) {
        for v in vs {
            self.vals.push(v.clone());
        }
    }

    /// Pop a whole result list — in REVERSE, the order they sit on the stack.
    fn pop_vals(&mut self, vs: &[Vt]) -> R<()> {
        for v in vs.iter().rev() {
            self.pop_expect(v)?;
        }
        Ok(())
    }

    fn push_ctrl(&mut self, op: &'static str, label: Option<String>, start: Vec<Vt>, end: Vec<Vt>) {
        let height = self.vals.len();
        self.ctrls.push(Ctrl {
            op,
            label,
            inited_entry: self.inited.clone(),
            start: start.clone(),
            end,
            height,
            unreachable: false,
        });
        self.push_vals(&start);
    }

    fn pop_ctrl(&mut self) -> R<Ctrl> {
        let (end, height) = {
            let f = self.ctrls.last().ok_or(Fail::Bail)?;
            (f.end.clone(), f.height)
        };
        self.pop_vals(&end)?;
        // ⛔ NOT `<`. Values LEFT OVER above the frame's floor are a type
        // mismatch just as surely as missing ones — that is the whole of
        // "type mismatch: values remaining on stack at end of block".
        if self.vals.len() != height {
            return tmis_d(format!(
                "values remaining on stack at end of block ({} above the frame)",
                self.vals.len() - height
            ));
        }
        self.ctrls.pop().ok_or(Fail::Bail)
    }

    fn unreachable(&mut self) -> R<()> {
        let h = self.ctrls.last().ok_or(Fail::Bail)?.height;
        self.vals.truncate(h);
        if let Some(f) = self.ctrls.last_mut() {
            f.unreachable = true;
        }
        Ok(())
    }

    /// A `loop` label takes the block's PARAMETERS (branching to it re-enters);
    /// every other frame takes its results.
    fn label_types(f: &Ctrl) -> &[Vt] {
        if f.op == "loop" { &f.start } else { &f.end }
    }

    /// Resolve a branch target: a depth, or a label NAME searched innermost-out.
    fn label_index(&self, s: &str) -> R<usize> {
        if let Some(name) = s.strip_prefix('$') {
            for (i, f) in self.ctrls.iter().rev().enumerate() {
                if f.label.as_deref() == Some(name) {
                    return Ok(i);
                }
            }
            // Unresolvable — `unknown label` is another rule's to report.
            return Err(Fail::Bail);
        }
        let n: usize = s.parse().map_err(|_| Fail::Bail)?;
        if n >= self.ctrls.len() {
            return Err(Fail::Bail);
        }
        Ok(n)
    }
}

// ── The fixed signatures ─────────────────────────────────────────────────────
//
// Every instruction whose operand and result types are the SAME whatever the
// module says. The parametric ones (`drop`, `select`), the indexed ones
// (`local.*`, `global.*`, `call`), the control ones and anything whose address
// width depends on a declared memory are NOT here — they are typed in `step`
// against the module's index spaces.
//
// Kept as plain type SPELLINGS rather than `Vt` on purpose: this table is a
// per-opcode fact with no module-relative content, so it stays movable to the
// shared opcode table if a second consumer (a validator over the emitted
// binary) ever appears. `Vt` carries `Heap::Concrete(usize)` — module type
// indices — and could not be expressed there.

/// The scalar type of one lane of a SIMD shape.
fn lane_scalar(shape: &str) -> Option<&'static str> {
    Some(match shape {
        "i8x16" | "i16x8" | "i32x4" => "i32",
        "i64x2" => "i64",
        "f32x4" => "f32",
        "f64x2" => "f64",
        _ => return None,
    })
}

fn static_num(t: &str) -> Option<&'static str> {
    Some(match t {
        "i32" => "i32",
        "i64" => "i64",
        "f32" => "f32",
        "f64" => "f64",
        "v128" => "v128",
        _ => return None,
    })
}

fn fixed_sig(name: &str) -> Option<(Vec<&'static str>, Vec<&'static str>)> {
    let (head, op) = name.split_once('.')?;

    // ── v128.* ───────────────────────────────────────────────────────────
    if head == "v128" {
        return Some(match op {
            "const" => (vec![], vec!["v128"]),
            "not" => (vec!["v128"], vec!["v128"]),
            "and" | "andnot" | "or" | "xor" => (vec!["v128", "v128"], vec!["v128"]),
            "bitselect" => (vec!["v128", "v128", "v128"], vec!["v128"]),
            "any_true" => (vec!["v128"], vec!["i32"]),
            _ => return None,
        });
    }

    // ── SIMD shapes ──────────────────────────────────────────────────────
    if let Some(scalar) = lane_scalar(head) {
        // Lane accessors carry a lane index immediate; the lane index rule is
        // a separate, already-landed check.
        if op == "splat" {
            return Some((vec![scalar], vec!["v128"]));
        }
        if op == "extract_lane" || op == "extract_lane_s" || op == "extract_lane_u" {
            return Some((vec!["v128"], vec![scalar]));
        }
        if op == "replace_lane" {
            return Some((vec!["v128", scalar], vec!["v128"]));
        }
        if op == "shuffle" {
            return Some((vec!["v128", "v128"], vec!["v128"]));
        }
        // A vector shift takes a SCALAR i32 shift count whatever the lane
        // width — `i64x2.shl` is [v128 i32] → [v128], not [v128 i64].
        if matches!(op, "shl" | "shr_s" | "shr_u") {
            return Some((vec!["v128", "i32"], vec!["v128"]));
        }
        if matches!(op, "all_true" | "bitmask") {
            return Some((vec!["v128"], vec!["i32"]));
        }
        // Unary lanewise.
        if op.starts_with("extend_low_")
            || op.starts_with("extend_high_")
            || op.starts_with("extadd_pairwise_")
            || op.starts_with("convert_")
            || op.starts_with("trunc_sat_")
            || op.starts_with("relaxed_trunc_")
            || matches!(
                op,
                "abs" | "neg" | "sqrt" | "ceil" | "floor" | "trunc" | "nearest" | "popcnt"
                    | "demote_f64x2_zero" | "promote_low_f32x4"
            )
        {
            return Some((vec!["v128"], vec!["v128"]));
        }
        // Everything else lanewise is binary v128 → v128: arithmetic, the
        // comparisons (which produce a MASK, not an i32), saturating and
        // widening forms, and the relaxed variants.
        if op.starts_with("add")
            || op.starts_with("sub")
            || op.starts_with("mul")
            || op.starts_with("min")
            || op.starts_with("max")
            || op.starts_with("avgr")
            || op.starts_with("narrow_")
            || op.starts_with("extmul_")
            || op.starts_with("dot_")
            || op.starts_with("q15mulr")
            || op.starts_with("relaxed_")
            || matches!(
                op,
                "div" | "pmin" | "pmax" | "swizzle"
                    | "eq" | "ne" | "lt" | "gt" | "le" | "ge"
                    | "lt_s" | "lt_u" | "gt_s" | "gt_u"
                    | "le_s" | "le_u" | "ge_s" | "ge_u"
            )
        {
            return Some((vec!["v128", "v128"], vec!["v128"]));
        }
        // `relaxed_laneselect` and friends take three vectors.
        if op == "relaxed_laneselect" {
            return Some((vec!["v128", "v128", "v128"], vec!["v128"]));
        }
        return None;
    }

    // ── scalar numeric ───────────────────────────────────────────────────
    let t = static_num(head)?;
    let is_int = t == "i32" || t == "i64";
    Some(match op {
        "const" => (vec![], vec![t]),
        // Integer unary and binary.
        "clz" | "ctz" | "popcnt" | "extend8_s" | "extend16_s" | "extend32_s" if is_int => {
            (vec![t], vec![t])
        }
        "eqz" if is_int => (vec![t], vec!["i32"]),
        "add" | "sub" | "mul" | "div_s" | "div_u" | "rem_s" | "rem_u" | "and" | "or" | "xor"
        | "shl" | "shr_s" | "shr_u" | "rotl" | "rotr"
            if is_int =>
        {
            (vec![t, t], vec![t])
        }
        "eq" | "ne" | "lt_s" | "lt_u" | "gt_s" | "gt_u" | "le_s" | "le_u" | "ge_s" | "ge_u"
            if is_int =>
        {
            (vec![t, t], vec!["i32"])
        }
        // Float unary and binary.
        "abs" | "neg" | "ceil" | "floor" | "trunc" | "nearest" | "sqrt" if !is_int => {
            (vec![t], vec![t])
        }
        "add" | "sub" | "mul" | "div" | "min" | "max" | "copysign" if !is_int => {
            (vec![t, t], vec![t])
        }
        "eq" | "ne" | "lt" | "gt" | "le" | "ge" if !is_int => (vec![t, t], vec!["i32"]),
        // Conversions. Each names its SOURCE type in the mnemonic, which is
        // what makes them checkable without a table entry apiece.
        _ => {
            let src = op
                .strip_prefix("wrap_")
                .or_else(|| op.strip_prefix("extend_"))
                .or_else(|| op.strip_prefix("trunc_"))
                .or_else(|| op.strip_prefix("trunc_sat_"))
                .or_else(|| op.strip_prefix("convert_"))
                .or_else(|| op.strip_prefix("demote_"))
                .or_else(|| op.strip_prefix("promote_"))
                .or_else(|| op.strip_prefix("reinterpret_"))?;
            // `i64.extend_i32_s` → `i32_s`; `f64.convert_i32_u` → `i32_u`.
            let src = src
                .strip_suffix("_s")
                .or_else(|| src.strip_suffix("_u"))
                .unwrap_or(src);
            // `i32.trunc_sat_f32_s` reaches here as `sat_f32` when the
            // `trunc_` prefix matched first; peel the qualifier.
            let src = src.strip_prefix("sat_").unwrap_or(src);
            (vec![static_num(src)?], vec![t])
        }
    })
}

/// The load/store family, whose ADDRESS operand type is the memory's, not a
/// constant. Returns `(value type, is_store)`.
fn mem_access_sig(name: &str) -> Option<(&'static str, bool)> {
    let (head, op) = name.split_once('.')?;
    if head == "v128" {
        return Some(match op {
            "load" | "load8x8_s" | "load8x8_u" | "load16x4_s" | "load16x4_u" | "load32x2_s"
            | "load32x2_u" | "load8_splat" | "load16_splat" | "load32_splat" | "load64_splat"
            | "load32_zero" | "load64_zero" => ("v128", false),
            "store" => ("v128", true),
            // The lane forms take the vector as a second operand.
            _ => return None,
        });
    }
    let t = static_num(head)?;
    if op == "load" || op == "store" {
        return Some((t, op == "store"));
    }
    for w in ["8", "16", "32"] {
        for sx in ["_s", "_u", ""] {
            if op == format!("load{w}{sx}") {
                return Some((t, false));
            }
        }
        if op == format!("store{w}") {
            return Some((t, true));
        }
    }
    None
}

// ── Building the module's index spaces ───────────────────────────────────────

/// The `param*`/`result*` of a `typeuse` or `func_type`, resolved.
fn sig_from_params(p: &Pair<Rule>, names: &HashMap<String, usize>) -> Option<FnSig> {
    let mut params = Vec::new();
    let mut results = Vec::new();
    for c in p.clone().into_inner() {
        match c.as_rule() {
            Rule::param => {
                for d in c.into_inner() {
                    if d.as_rule() != Rule::id {
                        params.push(parse_vt(d.as_str().trim(), names)?);
                    }
                }
            }
            Rule::result => {
                for d in c.into_inner() {
                    results.push(parse_vt(d.as_str().trim(), names)?);
                }
            }
            _ => {}
        }
    }
    Some(FnSig { params, results })
}

/// A `typeuse`: `(type $t)` alone, an inline signature, or both. When both are
/// written the inline one must agree, and it is the one the body is typed
/// against — so the inline spelling wins when it carries anything.
fn sig_from_typeuse(
    tu: &Pair<Rule>,
    ctx_type_sigs: &[Option<FnSig>],
    type_names: &HashMap<String, usize>,
    names: &HashMap<String, usize>,
) -> Option<(Option<usize>, FnSig)> {
    let written = tu.clone().into_inner().find(|c| c.as_rule() == Rule::index);
    let idx = written
        .as_ref()
        .and_then(|i| resolve_wast_index(i.as_str().trim(), type_names));
    // ⛔ WRITTEN-BUT-UNRESOLVABLE IS NOT "NO TYPE INDEX". The guard below keys
    // off `idx`, so a lookup MISS looked identical to a function that named no
    // type at all and quietly took the empty inline signature.
    if written.is_some() && idx.is_none() {
        return None;
    }
    // ⛔ THE TYPE INDEX IS AUTHORITATIVE WHEN IT RESOLVES. Where a function
    // writes both, the spec requires the inline spelling to MATCH the named
    // type — so they are interchangeable when legal, and the named one is the
    // only one that is complete when the inline half is partial.
    if let Some(i) = idx {
        if let Some(s) = ctx_type_sigs.get(i).cloned().flatten() {
            return Some((Some(i), s));
        }
    }
    let inline = sig_from_params(tu, names)?;
    // ⛔ AN UNRESOLVED `(type N)` IS NOT `() -> ()`. Falling through to the
    // empty inline signature GUESSES a type, and a function whose body then
    // leaves its real result on the stack reads as "values remaining on stack"
    // — a confident wrong answer on a VALID module, which is exactly what the
    // bail exists to prevent.
    if idx.is_some() && inline.params.is_empty() && inline.results.is_empty() {
        return None;
    }
    Some((idx, inline))
}

/// `$name` or a decimal index.
fn resolve_wast_index(s: &str, names: &HashMap<String, usize>) -> Option<usize> {
    let t = s.trim();
    match t.strip_prefix('$') {
        // ⛔ TWO KEY CONVENTIONS LIVE IN THIS FILE. `descriptor_type_table`
        // keys its name map by the id WITH the `$` still on it (`Rule::id`
        // matches `$name`), while `census_id`/`field_id` strip it. Looking up
        // only one silently missed EVERY `(type $t)` — and the miss was not
        // harmless: `sig_from_typeuse` fell through to the empty inline
        // signature and typed the function as `() -> ()`, which is a confident
        // wrong answer of exactly the kind the bail exists to prevent.
        Some(bare) => names.get(bare).or_else(|| names.get(t)).copied(),
        None => t.parse().ok(),
    }
}

/// The keyword an import descriptor or a field opens with.
fn head_keyword(p: &Pair<Rule>) -> String {
    let s = p.as_str().trim_start_matches('(').trim_start();
    s.split(|c: char| c.is_whitespace() || c == '(' || c == ')')
        .next()
        .unwrap_or("")
        .to_string()
}

/// The `id` immediately following a field's keyword, if it has one.
fn field_id(p: &Pair<Rule>) -> Option<String> {
    p.clone()
        .into_inner()
        .find(|c| c.as_rule() == Rule::id)
        .map(|c| c.as_str()[1..].to_string())
}

fn build_type_ctx(module: &Pair<Rule>) -> TypeCtx {
    let (types, type_names) = descriptor_type_table(module);
    let mut ctx = TypeCtx {
        types,
        type_names,
        ..Default::default()
    };
    let names = ctx.type_names.clone();

    // Pass 1 — the type section, so every later `(type $t)` resolves.
    for field in module_fields(module) {
        if field.as_rule() == Rule::type_field {
            let ft = find_rule(&field, Rule::func_type);
            ctx.type_sigs
                .push(ft.and_then(|f| sig_from_params(&f, &names)));
        }
    }

    // Pass 1b — THE IMPLICIT TYPES. §6.6.4: a `typeuse` written inline defines
    // a type when the module has no matching one, appended to the SAME index
    // space AFTER every explicit `(type …)`. func.wast pins the numbering
    // directly: one explicit `(type $t (func (param i32)))` is type 0, and the
    // inline `(func $f (result f64) …)` written ABOVE it becomes type 1.
    //
    // ⛔ DEDUPED, AND THE DEDUP IS WHAT MAKES THE INDICES RIGHT. `(func $g
    // (param i32))` in that same module adds nothing — it REUSES type 0 — so
    // appending unconditionally would shift every later index by one and
    // mistype the bodies that name them.
    let eligible = implicit_dedup_eligible(&ctx.types);
    let mut inline_sigs: Vec<Pair<Rule>> = Vec::new();
    collect_typeuses(module, &mut inline_sigs);
    for tu in inline_sigs {
        let has_index = tu.clone().into_inner().any(|c| c.as_rule() == Rule::index);
        if has_index {
            continue;
        }
        let Some(sig) = sig_from_params(&tu, &names) else { continue };
        // ⛔ `(func $f)` DEFINES THE TYPE `[] -> []` LIKE ANY OTHER SIGNATURE.
        // Skipping the empty one left such a function with no type index at
        // all, so `ref.func $f` bailed and the whole init expression went
        // unvalidated — an absent answer that looked like conservatism.
        let dup = ctx.type_sigs.iter().enumerate().any(|(i, t)| {
            eligible.get(i).copied().unwrap_or(true)
                && t.as_ref()
                    .is_some_and(|t| t.params == sig.params && t.results == sig.results)
        });
        if !dup {
            ctx.type_sigs.push(Some(sig));
        }
    }

    // Pass 2 — imports, THEN definitions. ⛔ An imported function occupies an
    // index before every defined one whatever the text order, so collecting in
    // document order across both spellings would shift every `call`.
    let mut defined: Vec<Pair<Rule>> = Vec::new();
    for field in module_fields(module) {
        match field.as_rule() {
            Rule::import_field => {
                if let Some(desc) = find_rule(&field, Rule::import_desc) {
                    add_entity(&mut ctx, &desc, &names);
                }
            }
            Rule::func_field
            | Rule::global_field
            | Rule::memory_field
            | Rule::table_field
            | Rule::tag_field => {
                // The inline spelling `(func $f (import "m" "n") …)` is an
                // import too, and belongs in the same early pass.
                if find_rule(&field, Rule::import_inline).is_some() {
                    add_entity(&mut ctx, &field, &names);
                } else {
                    defined.push(field);
                }
            }
            _ => {}
        }
    }
    // Everything in the global index space so far is an IMPORT, and that
    // boundary is the whole context a constant expression may read.
    ctx.imported_globals = ctx.globals.len();
    for field in defined {
        add_entity(&mut ctx, &field, &names);
    }

    // Pass 3 — element segments. They occupy their own index space, in
    // document order, and cannot be imported, so a single walk is the whole
    // rule. The `func`-index spelling `(elem $e func $f …)` has no written
    // reftype and means `funcref`.
    for field in module_fields(module) {
        if field.as_rule() != Rule::elem_field {
            continue;
        }
        let idx = ctx.elem_types.len();
        if let Some(id) = find_rule(&field, Rule::id) {
            ctx.elem_names.insert(id.as_str().to_string(), idx);
        }
        let vt = match find_rule(&field, Rule::ref_val_type) {
            Some(rt) => parse_vt(rt.as_str().trim(), &names),
            None => parse_vt("funcref", &names),
        };
        ctx.elem_types.push(vt);
    }
    ctx
}

/// Add one declared entity to its index space. Shared by both import
/// spellings and the defining fields, which is what keeps the two in step.
fn add_entity(ctx: &mut TypeCtx, p: &Pair<Rule>, names: &HashMap<String, usize>) {
    let kind = head_keyword(p);
    let id = field_id(p);
    match kind.as_str() {
        "func" => {
            let sig = find_rule(p, Rule::typeuse)
                .and_then(|tu| sig_from_typeuse(&tu, &ctx.type_sigs, &ctx.type_names, names));
            if let Some(n) = id {
                ctx.func_names.insert(n, ctx.funcs.len());
            }
            // ⛔ A FUNCTION DECLARED INLINE STILL HAS A TYPE INDEX — the
            // IMPLICIT one. `ref.func $f` is typed `(ref $t)` where `$t` is the
            // function's type, so taking only the WRITTEN `(type N)` left every
            // inline-signature function without one and made `ref.func` bail.
            // Pass 1b has already materialised those types, so the index is
            // recoverable by structural match — the same match the spec's own
            // dedup uses to decide whether to append in the first place.
            // ⛔ THE RECOVERY MUST USE THE SAME ELIGIBILITY AS PASS 1B'S DEDUP.
            // Pass 1b appends a fresh implicit type when the only structural
            // match sits inside a rec group; if this side still matched it, the
            // function would be handed the group member's index and the two
            // would agree again — the conflation restored one line later.
            let elig = implicit_dedup_eligible(&ctx.types);
            let ti = sig.as_ref().and_then(|(i, _)| *i).or_else(|| {
                let s = &sig.as_ref()?.1;
                ctx.type_sigs.iter().enumerate().position(|(i, t)| {
                    elig.get(i).copied().unwrap_or(true)
                        && t.as_ref()
                            .is_some_and(|t| t.params == s.params && t.results == s.results)
                })
            });
            ctx.func_types.push(ti);
            ctx.funcs.push(sig.map(|(_, s)| s));
        }
        "global" => {
            let vt = find_rule(p, Rule::global_type)
                .and_then(|g| {
                    let inner = g.as_str().trim();
                    let spelling = inner
                        .strip_prefix('(')
                        .and_then(|x| x.trim().strip_prefix("mut"))
                        .map(|x| x.trim().trim_end_matches(')').trim())
                        .unwrap_or(inner);
                    parse_vt(spelling, names)
                })
                ;
            let is_mut = find_rule(p, Rule::global_type)
                .is_some_and(|g| {
                    g.as_str()
                        .trim()
                        .strip_prefix('(')
                        .is_some_and(|x| x.trim().starts_with("mut"))
                });
            if let Some(n) = id {
                ctx.global_names.insert(n, ctx.globals.len());
            }
            ctx.globals.push(vt);
            ctx.global_mut.push(is_mut);
        }
        "memory" => {
            // ⛔ `addr_type` is a bare literal, so it never reaches the parse
            // tree as a pair — the memory's address width has to be read off
            // the text. It decides the operand type of every load and store
            // against this memory, and defaulting it to i32 mis-types every
            // memory64 body.
            let is64 = find_rule(p, Rule::mem_type)
                .map(|m| m.as_str().split_whitespace().next() == Some("i64"))
                .unwrap_or(false);
            ctx.mem_addr.push(if is64 { "i64" } else { "i32" });
        }
        "table" => {
            let tt = find_rule(p, Rule::table_type);
            let is64 = tt
                .as_ref()
                .map(|m| m.as_str().split_whitespace().next() == Some("i64"))
                .unwrap_or(false);
            let elem = tt
                .and_then(|t| {
                    t.into_inner()
                        .filter(|c| !matches!(c.as_rule(), Rule::integer))
                        .last()
                })
                .and_then(|c| parse_vt(c.as_str().trim(), names))
                ;
            if let Some(n) = id {
                ctx.table_names.insert(n, ctx.tables.len());
            }
            ctx.table_addr.push(if is64 { "i64" } else { "i32" });
            ctx.tables.push(elem);
        }
        "tag" => {
            // ⛔ A TAG'S SIGNATURE IS A `tag_type`, NOT A `typeuse`. This read
            // `Rule::typeuse` and found nothing, so every tag was recorded
            // with EMPTY params — `throw $e (i32.const 1)` then had nothing to
            // check against. `tag_type` is `(func param* result*)`,
            // `(type $t) param* result*`, or a bare `param* result*`, so the
            // `(type …)` spelling still has to go through the type table.
            let tt = find_rule(p, Rule::tag_type);
            let params = tt
                .as_ref()
                .and_then(|t| {
                    find_rule(t, Rule::index)
                        .and_then(|i| resolve_wast_index(i.as_str().trim(), &ctx.type_names))
                        .and_then(|i| ctx.type_sigs.get(i).cloned().flatten())
                        .map(|sig| sig.params)
                })
                .or_else(|| tt.as_ref().and_then(|t| {
                    sig_from_params(t, names).map(|s| s.params)
                }))
                .unwrap_or_default();
            if let Some(n) = id {
                ctx.tag_names.insert(n, ctx.tags.len());
            }
            ctx.tags.push(params);
        }
        _ => {}
    }
}

/// Every `typeuse` under a pair, in DOCUMENT ORDER — which is the order the
/// implicit types are appended in.
fn collect_typeuses<'a>(p: &Pair<'a, Rule>, out: &mut Vec<Pair<'a, Rule>>) {
    for c in p.clone().into_inner() {
        if c.as_rule() == Rule::typeuse {
            out.push(c.clone());
        }
        collect_typeuses(&c, out);
    }
}

/// The `module_field` children of a module, unwrapped one level.
fn module_fields<'a>(module: &Pair<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    module
        .clone()
        .into_inner()
        .filter_map(|p| {
            if p.as_rule() == Rule::module_field {
                p.into_inner().next()
            } else {
                Some(p)
            }
        })
        .collect()
}

/// The first descendant with the given rule, not descending past it.
fn find_rule<'a>(p: &Pair<'a, Rule>, want: Rule) -> Option<Pair<'a, Rule>> {
    for c in p.clone().into_inner() {
        if c.as_rule() == want {
            return Some(c);
        }
        if let Some(f) = find_rule(&c, want) {
            return Some(f);
        }
    }
    None
}

// ── Flattening: folded and plain reduce to ONE sequence ──────────────────────
//
// Plain form needs no work at all — `block`, `else` and `end` are already
// sibling `plain_instr`s carrying their `block_type` as an ordinary argument,
// so a linear walk types them as they come. Only the folded form has to be
// unfolded, and it unfolds POST-ORDER: a folded instruction's operands are
// evaluated before it.

fn flatten_instrs<'a>(seq: Vec<Pair<'a, Rule>>, out: &mut Vec<FlatInstr<'a>>) -> Option<()> {
    for p in seq {
        let p = if p.as_rule() == Rule::instr {
            p.into_inner().next()?
        } else {
            p
        };
        match p.as_rule() {
            Rule::plain_instr => {
                let (imms, ops) = split_immediates_and_operands(&p);
                // A plain instruction may still absorb folded operands.
                flatten_instrs(ops, out)?;
                let head = p
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::instr_name)?
                    .as_str()
                    .trim()
                    .to_string();
                let btypes: Vec<Pair<Rule>> = p
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .filter_map(|c| c.into_inner().next())
                    .filter(|c| c.as_rule() == Rule::block_type)
                    .collect();
                let label = p
                    .clone()
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::instr_arg)
                    .find_map(|c| c.into_inner().next())
                    .filter(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
                let imms = imms
                    .into_iter()
                    .filter(|s| !s.starts_with("(type") && !s.starts_with("(result")
                        && !s.starts_with("(param"))
                    .collect();
                out.push(FlatInstr { head, imms, btypes, label, clauses: vec![] });
            }
            Rule::folded_instr => {
                let head = head_keyword(&p);
                // ⛔ TWO PLACES, ONE MEANING. `block`/`loop`/`if` name
                // `block_type*` directly in the grammar, but on every other
                // folded instruction the same `(type $t)` arrives inside an
                // `instr_arg` — so `(call_indirect (type $t) …)` collected
                // ZERO signatures and typed as `[] -> []`, which is a wrong
                // answer rather than an absent one.
                let btypes: Vec<Pair<Rule>> = p
                    .clone()
                    .into_inner()
                    .filter_map(|c| match c.as_rule() {
                        Rule::block_type => Some(c),
                        Rule::instr_arg => c
                            .into_inner()
                            .next()
                            .filter(|i| i.as_rule() == Rule::block_type),
                        _ => None,
                    })
                    .collect();
                let label = p
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::id)
                    .map(|c| c.as_str()[1..].to_string());
                match head.as_str() {
                    "block" | "loop" => {
                        let body: Vec<Pair<Rule>> = p
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::instr)
                            .collect();
                        out.push(FlatInstr {
                            head,
                            imms: vec![],
                            btypes,
                            label,
                            clauses: vec![],
                        });
                        flatten_instrs(body, out)?;
                        out.push(FlatInstr {
                            head: "end".into(),
                            imms: vec![],
                            btypes: vec![],
                            label: None,
                            clauses: vec![],
                        });
                    }
                    "if" => {
                        // ⛔ THE CONDITION IS EVALUATED BEFORE THE FRAME OPENS.
                        // In `(if (cond) (then …))` the `instr` children are
                        // the condition and nothing else — `then`/`else` are
                        // their own rules — so they flatten AHEAD of the `if`.
                        let cond: Vec<Pair<Rule>> = p
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::instr)
                            .collect();
                        flatten_instrs(cond, out)?;
                        out.push(FlatInstr {
                            head: "if".into(),
                            imms: vec![],
                            btypes,
                            label,
                            clauses: vec![],
                        });
                        for c in p.clone().into_inner() {
                            match c.as_rule() {
                                Rule::then_block => {
                                    let b: Vec<Pair<Rule>> =
                                        c.into_inner().filter(|x| x.as_rule() == Rule::instr).collect();
                                    flatten_instrs(b, out)?;
                                }
                                Rule::else_block => {
                                    out.push(FlatInstr {
                                        head: "else".into(),
                                        imms: vec![],
                                        btypes: vec![],
                                        label: None,
                                        clauses: vec![],
                                    });
                                    let b: Vec<Pair<Rule>> =
                                        c.into_inner().filter(|x| x.as_rule() == Rule::instr).collect();
                                    flatten_instrs(b, out)?;
                                }
                                _ => {}
                            }
                        }
                        out.push(FlatInstr {
                            head: "end".into(),
                            imms: vec![],
                            btypes: vec![],
                            label: None,
                            clauses: vec![],
                        });
                    }
                    // ⛔ `try_table` OPENS AN ORDINARY BLOCK. Bailing here
                    // abandoned the whole function, so every one of
                    // `try_table.wast`'s nine assertions went undetected — and
                    // two of them (`(try_table (result i32))` with an empty
                    // body) are plain block-arity failures that need no
                    // exception machinery at all. The handlers are branch
                    // targets, not frame kinds: the clauses ride along and the
                    // body types against a normal block frame.
                    //
                    // The legacy `try`/`catch` shape is a different instruction
                    // with a different frame; it still abstains.
                    "try_table" => {
                        let clauses: Vec<Pair<Rule>> = p
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::try_clause)
                            .collect();
                        let body: Vec<Pair<Rule>> = p
                            .clone()
                            .into_inner()
                            .filter(|c| c.as_rule() == Rule::instr)
                            .collect();
                        out.push(FlatInstr {
                            head: "try_table".into(),
                            imms: vec![],
                            btypes,
                            label,
                            clauses,
                        });
                        flatten_instrs(body, out)?;
                        out.push(FlatInstr {
                            head: "end".into(),
                            imms: vec![],
                            btypes: vec![],
                            label: None,
                            clauses: vec![],
                        });
                    }
                    "try" => return None,
                    _ => {
                        let (imms, ops) = split_immediates_and_operands(&p);
                        flatten_instrs(ops, out)?;
                        let imms = imms
                            .into_iter()
                            .filter(|s| !s.starts_with("(type") && !s.starts_with("(result")
                                && !s.starts_with("(param"))
                            .collect();
                        out.push(FlatInstr { head, imms, btypes, label, clauses: vec![] });
                    }
                }
            }
            _ => {}
        }
    }
    Some(())
}

// ── The typing rules ─────────────────────────────────────────────────────────

impl<'a> Tv<'a> {
    fn num(t: &str) -> Vt {
        Vt::Num(static_num(t).unwrap_or("i32"))
    }

    /// A block signature: `(type $t)`, or inline `(param …)`/`(result …)`, or
    /// nothing at all.
    fn block_sig(&self, btypes: &[Pair<Rule>]) -> R<(Vec<Vt>, Vec<Vt>)> {
        let mut params = Vec::new();
        let mut results = Vec::new();
        for b in btypes {
            let text = b.as_str().trim();
            let body = text.trim_start_matches('(').trim();
            if let Some(rest) = body.strip_prefix("type") {
                let name = rest.trim().trim_end_matches(')').trim();
                let i = resolve_wast_index(name, &self.ctx.type_names).ok_or(Fail::Bail)?;
                let sig = self.ctx.type_sigs.get(i).cloned().flatten().ok_or(Fail::Bail)?;
                params.extend(sig.params);
                results.extend(sig.results);
            } else if body.starts_with("result") || body.starts_with("param") {
                let is_param = body.starts_with("param");
                for c in b.clone().into_inner() {
                    if c.as_rule() == Rule::id {
                        continue;
                    }
                    let v = parse_vt(c.as_str().trim(), &self.ctx.type_names)
                        .ok_or(Fail::Bail)?;
                    if is_param {
                        params.push(v)
                    } else {
                        results.push(v)
                    }
                }
            }
        }
        Ok((params, results))
    }

    /// The address type of the memory an access names — `i32`, or `i64` when
    /// the memory was declared `i64`.
    fn mem_addr(&self, imms: &[String]) -> R<Vt> {
        let idx = imms
            .iter()
            .find(|s| !s.starts_with("offset=") && !s.starts_with("align="))
            .and_then(|s| resolve_wast_index(s, &HashMap::new()))
            .unwrap_or(0);
        let t = self
            .ctx
            .mem_addr
            .get(idx)
            .or_else(|| self.ctx.mem_addr.first())
            .ok_or(Fail::Bail)?;
        Ok(Vt::Num(t))
    }

    fn table_of(&self, imms: &[String]) -> R<(usize, Vt, Vt)> {
        let idx = imms
            .first()
            .and_then(|s| resolve_wast_index(s, &self.ctx.table_names))
            .unwrap_or(0);
        // A table whose element type did not parse ABSTAINS. It must not
        // fall through to a bottom that accepts anything.
        let elem = self.ctx.tables.get(idx).cloned().flatten().ok_or(Fail::Bail)?;
        let addr = Vt::Num(self.ctx.table_addr.get(idx).copied().ok_or(Fail::Bail)?);
        Ok((idx, elem, addr))
    }

    fn local_at(&self, s: &str) -> R<Vt> {
        let s = s.trim();
        let i = match s.strip_prefix('$') {
            // ⛔ THE SAME TWO-CONVENTION SPLIT AS `resolve_wast_index`, AND
            // HERE IT WAS SILENT. `func_locals` keys its map by the id WITH
            // the `$` still attached; stripping before the lookup missed every
            // NAMED local, so any function written with `(param $x …)` bailed
            // and went completely unvalidated. It looked like conservatism —
            // the bail reports nothing — which is exactly why nothing failed:
            // 8 simd lane files asserted "type mismatch" against a pass that
            // had quietly declined to type the function at all.
            Some(bare) => *self
                .local_names
                .get(bare)
                .or_else(|| self.local_names.get(s))
                .ok_or(Fail::Bail)?,
            None => s.parse().map_err(|_| Fail::Bail)?,
        };
        self.locals.get(i).cloned().ok_or(Fail::Bail)
    }

    /// The INDEX behind a local's spelling — same two-convention resolution as
    /// `local_at`, which returns only the type.
    fn local_index(&self, s: &str) -> R<usize> {
        let s = s.trim();
        match s.strip_prefix('$') {
            Some(bare) => self
                .local_names
                .get(bare)
                .or_else(|| self.local_names.get(s))
                .copied()
                .ok_or(Fail::Bail),
            None => s.parse().map_err(|_| Fail::Bail),
        }
    }

    /// §3.4.1: reading a local that is not yet initialized on every path.
    ///
    /// ⛔ SKIPPED IN UNREACHABLE CODE. After `unreachable` the stack is
    /// polymorphic and the path is dead; enforcing initialization there would
    /// reject modules the spec accepts, and this rule runs over every module
    /// that compiles, not only the ones asserted invalid.
    fn read_local(&mut self, s: &str) -> R<()> {
        let i = self.local_index(s)?;
        let dead = self.ctrls.last().is_some_and(|f| f.unreachable);
        if !dead && !self.inited.get(i).copied().unwrap_or(true) {
            return Err(Fail::Mismatch("uninitialized local".to_string()));
        }
        Ok(())
    }

    fn write_local(&mut self, s: &str) -> R<()> {
        let i = self.local_index(s)?;
        if let Some(f) = self.inited.get_mut(i) {
            *f = true;
        }
        Ok(())
    }

    /// The type index an instruction's first immediate names.
    fn type_index(&self, imm: Option<&String>) -> R<usize> {
        resolve_wast_index(imm.ok_or(Fail::Bail)?, &self.ctx.type_names).ok_or(Fail::Bail)
    }

    /// A struct's field types, in declaration order.
    fn field_types(&self, t: &DescType) -> R<Vec<Vt>> {
        t.field_types
            .iter()
            .map(|f| {
                let bare = strip_mut(f);
                if matches!(bare.as_str(), "i8" | "i16") {
                    // Packed storage is written as i32.
                    Ok(Vt::Num("i32"))
                } else {
                    parse_vt(&bare, &self.ctx.type_names).ok_or(Fail::Bail)
                }
            })
            .collect()
    }

    /// An array's element type as a value type (packed storage reads as i32).
    fn array_elem_vt(&self, t: &DescType) -> R<Vt> {
        let raw = t.array_elem.clone().ok_or(Fail::Bail)?;
        if matches!(raw.as_str(), "i8" | "i16") {
            return Ok(Vt::Num("i32"));
        }
        parse_vt(&raw, &self.ctx.type_names).ok_or(Fail::Bail)
    }

    fn is_ref(v: &Vt) -> bool {
        matches!(v, Vt::Ref(_) | Vt::Bottom)
    }

    fn step(&mut self, ins: &FlatInstr, func_results: &[Vt]) -> R<()> {
        // The walker's `@@` channels ride in the name; the mnemonic is what
        // precedes the first one.
        let head = ins.head.split("@@").next().unwrap_or(&ins.head).to_string();
        let name = head.as_str();
        let imms = &ins.imms;

        // §3.4.10: a constant expression admits only the constant
        // instructions. `(global i32 (i32.ctz (i32.const 0)))` types
        // perfectly and is still invalid, so this cannot fall out of the
        // typing rules — it is a separate, syntactic admissibility check, and
        // it runs only where `global_limit` says we are in an init expression.
        if self.global_limit.is_some() && !Self::is_const_instr(name) {
            return Err(Fail::Mismatch("constant expression required".to_string()));
        }

        match name {
            "nop" => Ok(()),
            "unreachable" => self.unreachable(),
            "drop" => {
                self.pop_val()?;
                Ok(())
            }
            "select" => {
                self.pop_expect(&Vt::Num("i32"))?;
                if !ins.btypes.is_empty() {
                    let (_, res) = self.block_sig(&ins.btypes)?;
                    // ⛔ A TYPED `select` TAKES EXACTLY ONE RESULT, AND THE
                    // SPEC GIVES THAT ITS OWN DIAGNOSTIC. `(select (result))`
                    // and `(select (result i32 i32))` assert "invalid result
                    // arity", not "type mismatch" — reporting the generic
                    // message there fails the fixture even though the module
                    // is correctly rejected.
                    if res.len() != 1 {
                        return Err(Fail::Mismatch("invalid result arity".to_string()));
                    }
                    let t = res.first().cloned().ok_or(Fail::Bail)?;
                    self.pop_expect(&t)?;
                    self.pop_expect(&t)?;
                    self.push_val(t);
                    return Ok(());
                }
                let t1 = self.pop_val()?;
                let t2 = self.pop_val()?;
                // ⛔ UNTYPED `select` IS NUMERIC/VECTOR ONLY. A reference
                // operand needs the `(result t)` spelling, and the suite
                // asserts exactly that.
                if Self::is_ref(&t1) && !matches!(t1, Vt::Bottom) {
                    return tmis();
                }
                if Self::is_ref(&t2) && !matches!(t2, Vt::Bottom) {
                    return tmis();
                }
                let t = match (&t1, &t2) {
                    (Vt::Bottom, x) | (x, Vt::Bottom) => x.clone(),
                    (a, b) if a == b => a.clone(),
                    _ => return tmis(),
                };
                self.push_val(t);
                Ok(())
            }

            // ── control ──────────────────────────────────────────────────
            "block" | "loop" => {
                let (p, r) = self.block_sig(&ins.btypes)?;
                self.pop_vals(&p)?;
                let op = if name == "loop" { "loop" } else { "block" };
                self.push_ctrl(op, ins.label.clone(), p, r);
                Ok(())
            }
            "if" => {
                let (p, r) = self.block_sig(&ins.btypes)?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_vals(&p)?;
                self.push_ctrl("if", ins.label.clone(), p, r);
                Ok(())
            }
            // §3.4.8.19: `try_table bt (catch …)* instr*` is an ordinary block
            // whose handlers are BRANCHES. Each clause carries a fixed list to
            // its target label: a tag's parameters, `catch_ref` those plus the
            // caught `exnref`, `catch_all` nothing, `catch_all_ref` the exnref.
            //
            // ⛔ THE CLAUSE LABELS RESOLVE IN THE *OUTER* CONTEXT — checked
            // BEFORE the frame is pushed. Fixture #3 is the one that decides
            // it: `(func (result exnref) (try_table (catch 0 0)) (unreachable))`
            // is invalid only if label 0 is the FUNCTION, whose result is
            // `exnref`, against a `catch` that carries nothing. Resolving it
            // against the try_table's own frame — whose result is empty — makes
            // that module type clean.
            "try_table" => {
                let (p, r) = self.block_sig(&ins.btypes)?;
                let exn = parse_vt("exnref", &self.ctx.type_names).ok_or(Fail::Bail)?;
                for cl in &ins.clauses {
                    let kind = head_keyword(cl);
                    let ids: Vec<String> = cl
                        .clone()
                        .into_inner()
                        .filter(|c| c.as_rule() == Rule::index)
                        .map(|c| c.as_str().trim().to_string())
                        .collect();
                    let (tag_key, label_key) = match kind.as_str() {
                        "catch" | "catch_ref" => (ids.first(), ids.get(1)),
                        _ => (None, ids.first()),
                    };
                    let mut carried: Vec<Vt> = Vec::new();
                    if let Some(k) = tag_key {
                        let ti = resolve_wast_index(k, &self.ctx.tag_names)
                            .filter(|i| *i < self.ctx.tags.len())
                            .ok_or_else(|| Fail::Mismatch(unknown_index_msg("tag", k)))?;
                        carried.extend(self.ctx.tags[ti].iter().cloned());
                    }
                    if matches!(kind.as_str(), "catch_ref" | "catch_all_ref") {
                        carried.push(exn.clone());
                    }
                    let l = self.label_index(label_key.ok_or(Fail::Bail)?)?;
                    let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                    let agrees = lt.len() == carried.len()
                        && carried
                            .iter()
                            .zip(lt.iter())
                            .all(|(g, w)| vt_subtype(g, w, &self.ctx.types));
                    if !agrees {
                        let show = |vs: &[Vt]| {
                            vs.iter().map(vt_show).collect::<Vec<_>>().join(" ")
                        };
                        return tmis_d(format!(
                            "handler carries [{}] but label takes [{}]",
                            show(&carried),
                            show(&lt)
                        ));
                    }
                }
                self.pop_vals(&p)?;
                self.push_ctrl("block", ins.label.clone(), p, r);
                Ok(())
            }
            "else" => {
                let f = self.pop_ctrl()?;
                if f.op != "if" {
                    return tmis();
                }
                // The `else` arm starts from the state at the `if`, not from
                // whatever the `then` arm initialized — `local_init.wast`'s
                // uninit-in-else fixture is exactly that module.
                self.inited = f.inited_entry.clone();
                self.push_ctrl("else", f.label.clone(), f.start, f.end);
                Ok(())
            }
            "end" => {
                let f = self.pop_ctrl()?;
                self.inited = f.inited_entry.clone();
                // ⛔ AN `if` WITH NO `else` IS THE IDENTITY ON ITS PARAMETERS,
                // so it is only well typed when they ARE its results. Without
                // this an `(if (result i32) (then …))` types clean.
                if f.op == "if" && f.start != f.end {
                    return tmis_d(
                        "if without else must have the same parameters and results".to_string(),
                    );
                }
                self.push_vals(&f.end);
                Ok(())
            }
            // §3.3.8: `throw x` pops the tag's parameters and is then
            // unreachable — it never falls through, so the rest of the block
            // types against the polymorphic stack.
            "throw" => {
                let key = imms.first().ok_or(Fail::Bail)?;
                let i = resolve_wast_index(key, &self.ctx.tag_names)
                    .filter(|i| *i < self.ctx.tags.len())
                    .ok_or_else(|| Fail::Mismatch(unknown_index_msg("tag", key)))?;
                let params = self.ctx.tags[i].clone();
                // ⛔ THE SUITE'S ONLY TWO DETAILED `type mismatch` WORDINGS
                // ARE BOTH HERE, and the comparison is one-directional — our
                // reason must CONTAIN the fixture's text, so the generic
                // "type mismatch: stack underflow in block" does NOT discharge
                // `"type mismatch: instruction requires [i32] but stack has []"`.
                //
                // Skipped once the frame is unreachable: the polymorphic
                // bottom legitimately satisfies any requirement there, and
                // reporting a mismatch would reject valid code after a `br`.
                let f = self.ctrls.last().ok_or(Fail::Bail)?;
                if !f.unreachable {
                    let have: Vec<Vt> = self.vals[f.height..].to_vec();
                    let disagrees = have.len() < params.len()
                        || !params.iter().rev().zip(have.iter().rev()).all(|(want, got)| {
                            matches!(got, Vt::Bottom)
                                || matches!(want, Vt::Bottom)
                                || vt_subtype(got, want, &self.ctx.types)
                        });
                    if disagrees {
                        let show = |vs: &[Vt]| {
                            vs.iter().map(vt_show).collect::<Vec<_>>().join(" ")
                        };
                        return Err(Fail::Mismatch(format!(
                            "type mismatch: instruction requires [{}] but stack has [{}]",
                            show(&params),
                            show(&have)
                        )));
                    }
                }
                self.pop_vals(&params)?;
                self.unreachable()
            }
            // §3.4.8.6: `throw_ref` re-raises an exception it is HANDED, so it
            // pops one `exnref` and is stack-polymorphic after — exactly the
            // shape of `throw`, minus the tag.
            //
            // ⛔ THE OPERAND IS THE WHOLE RULE. Both fixtures spell it with an
            // EMPTY stack (`(func (throw_ref))`, `(func (block (throw_ref))))`,
            // so an arm that only went unreachable would discharge neither: the
            // pop is what rejects them, and going unreachable first would eat
            // the very mismatch being asserted.
            "throw_ref" => {
                let exn = parse_vt("exnref", &self.ctx.type_names).ok_or(Fail::Bail)?;
                self.pop_expect(&exn)?;
                self.unreachable()
            }
            "br" => {
                let l = self.label_index(imms.first().ok_or(Fail::Bail)?)?;
                let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                self.pop_vals(&lt)?;
                self.unreachable()
            }
            "br_if" => {
                self.pop_expect(&Vt::Num("i32"))?;
                let l = self.label_index(imms.first().ok_or(Fail::Bail)?)?;
                let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                self.pop_vals(&lt)?;
                self.push_vals(&lt);
                Ok(())
            }
            "br_table" => {
                self.pop_expect(&Vt::Num("i32"))?;
                if imms.is_empty() {
                    return Err(Fail::Bail);
                }
                // The last label is the default; every arm must agree with it.
                let dl = self.label_index(imms.last().ok_or(Fail::Bail)?)?;
                let dt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - dl]).to_vec();
                for s in &imms[..imms.len() - 1] {
                    let l = self.label_index(s)?;
                    let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                    if lt.len() != dt.len() {
                        return tmis();
                    }
                }
                self.pop_vals(&dt)?;
                self.unreachable()
            }
            "return" => {
                let r = func_results.to_vec();
                self.pop_vals(&r)?;
                self.unreachable()
            }

            // ── calls ────────────────────────────────────────────────────
            "call" | "return_call" => {
                let i = resolve_wast_index(imms.first().ok_or(Fail::Bail)?, &self.ctx.func_names)
                    .ok_or(Fail::Bail)?;
                let sig = self.ctx.funcs.get(i).cloned().flatten().ok_or(Fail::Bail)?;
                self.pop_vals(&sig.params)?;
                if name == "return_call" {
                    if sig.results != func_results {
                        return tmis();
                    }
                    return self.unreachable();
                }
                self.push_vals(&sig.results);
                Ok(())
            }
            "call_indirect" | "return_call_indirect" => {
                let (p, r) = self.block_sig(&ins.btypes)?;
                let (_, _, addr) = self.table_of(imms)?;
                self.pop_expect(&addr)?;
                self.pop_vals(&p)?;
                if name == "return_call_indirect" {
                    if r != func_results {
                        return tmis();
                    }
                    return self.unreachable();
                }
                self.push_vals(&r);
                Ok(())
            }
            "call_ref" | "return_call_ref" => {
                let ti = resolve_wast_index(imms.first().ok_or(Fail::Bail)?, &self.ctx.type_names)
                    .ok_or(Fail::Bail)?;
                let sig = self.ctx.type_sigs.get(ti).cloned().flatten().ok_or(Fail::Bail)?;
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(ti),
                }))?;
                self.pop_vals(&sig.params)?;
                if name == "return_call_ref" {
                    if sig.results != func_results {
                        return tmis();
                    }
                    return self.unreachable();
                }
                self.push_vals(&sig.results);
                Ok(())
            }

            // ── variables ────────────────────────────────────────────────
            "local.get" => {
                let key = imms.first().ok_or(Fail::Bail)?;
                let t = self.local_at(key)?;
                self.read_local(key)?;
                self.push_val(t);
                Ok(())
            }
            "local.set" => {
                let key = imms.first().ok_or(Fail::Bail)?;
                let t = self.local_at(key)?;
                self.pop_expect(&t)?;
                self.write_local(key)?;
                Ok(())
            }
            "local.tee" => {
                // ⛔ A `tee` WRITES BEFORE IT READS BACK. Its result is the
                // value it just stored, so it INITIALIZES the local — checking
                // it as a read would reject `(local.tee $x v)` on a local that
                // this very instruction makes valid.
                let key = imms.first().ok_or(Fail::Bail)?;
                let t = self.local_at(key)?;
                self.pop_expect(&t)?;
                self.write_local(key)?;
                self.push_val(t);
                Ok(())
            }
            "global.get" | "global.set" => {
                let key = imms.first().ok_or(Fail::Bail)?;
                // ⛔ ONE `None`, TWO MEANINGS. An index that does not resolve
                // used to bail — indistinguishable from "this spelling was not
                // understood" — so `(global.get 1)` in a one-global module
                // proved nothing and every "unknown global" fixture went
                // unanswered. Out of range is a DIAGNOSIS, not an abstention.
                let i = resolve_wast_index(key, &self.ctx.global_names)
                    .filter(|i| *i < self.ctx.globals.len())
                    .ok_or_else(|| Fail::Mismatch(unknown_index_msg("global", key)))?;
                if self.global_limit.is_some_and(|n| i >= n) {
                    return Err(Fail::Mismatch(unknown_index_msg("global", key)));
                }
                // A constant expression may read an IMMUTABLE global only, and
                // the suite gives a mutable one the constant-expression
                // wording rather than "immutable global" — the same import is
                // perfectly legal to read from a function body.
                if self.global_limit.is_some()
                    && self.ctx.global_mut.get(i).copied().unwrap_or(false)
                {
                    return Err(Fail::Mismatch("constant expression required".to_string()));
                }
                let t = self.ctx.globals.get(i).cloned().flatten().ok_or(Fail::Bail)?;
                if name == "global.get" {
                    self.push_val(t);
                } else {
                    // §3.3.5: the target must be mutable. Checked BEFORE the
                    // operand is popped — a `global.set` on an immutable
                    // global whose operand also disagrees asserts
                    // "immutable global", not "type mismatch".
                    if !self.ctx.global_mut.get(i).copied().unwrap_or(true) {
                        return Err(Fail::Mismatch("immutable global".to_string()));
                    }
                    self.pop_expect(&t)?;
                }
                Ok(())
            }

            // ── references ───────────────────────────────────────────────
            "ref.null" => {
                let a = imms.first().ok_or(Fail::Bail)?;
                let heap = match abs_heap(a) {
                    Some(x) => Heap::Abs(x),
                    None => Heap::Concrete(
                        resolve_wast_index(a, &self.ctx.type_names).ok_or(Fail::Bail)?,
                    ),
                };
                self.push_val(Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap,
                }));
                Ok(())
            }
            "ref.is_null" => {
                let t = self.pop_val()?;
                if !Self::is_ref(&t) {
                    return tmis();
                }
                self.push_val(Vt::Num("i32"));
                Ok(())
            }
            "ref.as_non_null" => {
                let t = self.pop_val()?;
                match t {
                    Vt::Bottom => self.push_val(Vt::Bottom),
                    Vt::Ref(r) => self.push_val(Vt::Ref(RefT {
                        nullable: false,
                        ..r
                    })),
                    _ => return tmis(),
                }
                Ok(())
            }
            "ref.func" => {
                let f = resolve_wast_index(imms.first().ok_or(Fail::Bail)?, &self.ctx.func_names)
                    .ok_or(Fail::Bail)?;
                // Without the function's TYPE index this cannot be given the
                // `(ref $t)` the spec assigns it, and approximating it as
                // `(ref func)` would fail a perfectly good `(ref $t)` context.
                let ti = self.ctx.func_types.get(f).copied().flatten().ok_or(Fail::Bail)?;
                self.push_val(Vt::Ref(RefT {
                    nullable: false,
                    exact: false,
                    heap: Heap::Concrete(ti),
                }));
                Ok(())
            }
            "br_on_null" | "br_on_non_null" => {
                let l = self.label_index(imms.first().ok_or(Fail::Bail)?)?;
                let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                let t = self.pop_val()?;
                let nn = match &t {
                    Vt::Bottom => Vt::Bottom,
                    Vt::Ref(r) => Vt::Ref(RefT {
                        nullable: false,
                        ..r.clone()
                    }),
                    _ => return tmis(),
                };
                if name == "br_on_null" {
                    self.pop_vals(&lt)?;
                    self.push_vals(&lt);
                    self.push_val(nn);
                } else {
                    // Branches carrying the non-null reference; falls through
                    // with it consumed.
                    let mut want = lt.clone();
                    if want.pop().is_none() {
                        return tmis();
                    }
                    self.pop_vals(&want)?;
                    self.push_vals(&want);
                }
                Ok(())
            }
            // §3.4.8.10-11: `br_on_cast l rt1 rt2` takes an `rt1`, branches to
            // `l` carrying `rt2` when the cast succeeds, and falls through with
            // the DIFFERENCE `rt1\rt2`. `br_on_cast_fail` swaps those two: it
            // branches with the difference and falls through with `rt2`.
            //
            // ⛔ `rt2 <: rt1` IS A SIDE CONDITION, NOT A CONSEQUENCE. Three of
            // the twelve fixtures are invalid for that reason alone — casting
            // `eqref` to `anyref` widens, and no stack shape reveals it — so it
            // has to be asserted before any operand is touched.
            "br_on_cast" | "br_on_cast_fail" => {
                let l = self.label_index(imms.first().ok_or(Fail::Bail)?)?;
                let rt1 = parse_vt(imms.get(1).ok_or(Fail::Bail)?, &self.ctx.type_names)
                    .ok_or(Fail::Bail)?;
                let rt2 = parse_vt(imms.get(2).ok_or(Fail::Bail)?, &self.ctx.type_names)
                    .ok_or(Fail::Bail)?;
                let (Vt::Ref(r1), Vt::Ref(r2)) = (&rt1, &rt2) else {
                    return tmis_d("br_on_cast needs two reference types".to_string());
                };
                if !vt_subtype(&rt2, &rt1, &self.ctx.types) {
                    return tmis_d(format!(
                        "{} is not a subtype of {}",
                        vt_show(&rt2),
                        vt_show(&rt1)
                    ));
                }
                // §4.2.9.2: `(ref null1 ht1) \ (ref null2 ht2)` keeps ht1 and
                // loses nullability exactly when rt2 could have absorbed the
                // null — i.e. when rt2 is itself nullable.
                let diff = Vt::Ref(RefT {
                    nullable: r1.nullable && !r2.nullable,
                    ..r1.clone()
                });
                let (carried, falls) = if name == "br_on_cast" {
                    (rt2.clone(), diff)
                } else {
                    (diff, rt2.clone())
                };
                let lt = Self::label_types(&self.ctrls[self.ctrls.len() - 1 - l]).to_vec();
                let mut want = lt.clone();
                let slot = want
                    .pop()
                    .ok_or_else(|| Fail::Mismatch("type mismatch: label takes no value".into()))?;
                if !vt_subtype(&carried, &slot, &self.ctx.types) {
                    return tmis_d(format!(
                        "branch carries {} but label takes {}",
                        vt_show(&carried),
                        vt_show(&slot)
                    ));
                }
                // The operand sits ON TOP of the `t*` the label also takes, so
                // it comes off first; `t*` stays put on the fall-through path.
                self.pop_expect(&rt1)?;
                self.pop_vals(&want)?;
                self.push_vals(&want);
                self.push_val(falls);
                Ok(())
            }

            // ── memory ───────────────────────────────────────────────────
            "memory.size" => {
                let a = self.mem_addr(imms)?;
                self.push_val(a);
                Ok(())
            }
            "memory.grow" => {
                let a = self.mem_addr(imms)?;
                self.pop_expect(&a)?;
                self.push_val(a);
                Ok(())
            }
            "memory.fill" => {
                let a = self.mem_addr(imms)?;
                self.pop_expect(&a)?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&a)?;
                Ok(())
            }
            "memory.copy" => {
                let a = self.mem_addr(imms)?;
                self.pop_expect(&a)?;
                self.pop_expect(&a)?;
                self.pop_expect(&a)?;
                Ok(())
            }
            "memory.init" => {
                // ⛔ THE MEMIDX PEEL AGAIN. `memory.init $d` names a DATA
                // segment; only the multi-memory spelling puts a memidx first,
                // and guessing wrong silently types the address against
                // ANOTHER memory's width. Memory 0 is the one form both
                // spellings agree on.
                let a = self.mem_addr(&[])?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&a)?;
                Ok(())
            }
            "data.drop" | "elem.drop" => Ok(()),

            // ── tables ───────────────────────────────────────────────────
            "table.get" => {
                let (_, e, a) = self.table_of(imms)?;
                self.pop_expect(&a)?;
                self.push_val(e);
                Ok(())
            }
            "table.set" => {
                let (_, e, a) = self.table_of(imms)?;
                self.pop_expect(&e)?;
                self.pop_expect(&a)?;
                Ok(())
            }
            "table.size" => {
                let (_, _, a) = self.table_of(imms)?;
                self.push_val(a);
                Ok(())
            }
            "table.grow" => {
                let (_, e, a) = self.table_of(imms)?;
                self.pop_expect(&a)?;
                self.pop_expect(&e)?;
                self.push_val(a);
                Ok(())
            }
            "table.fill" => {
                let (_, e, a) = self.table_of(imms)?;
                self.pop_expect(&a)?;
                self.pop_expect(&e)?;
                self.pop_expect(&a)?;
                Ok(())
            }
            "table.copy" => {
                // ⛔ THE TWO TABLES MAY HAVE DIFFERENT ADDRESS WIDTHS.
                // `table.copy $t32 $t64` takes an i32 destination index and an
                // i64 source index; collapsing both to table 0 typed the source
                // against the wrong width. With TWO immediates written the
                // positions are unambiguous — dest first, then source — and
                // the length is the NARROWER of the two.
                let ids: Vec<String> = imms
                    .iter()
                    .filter(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                    .cloned()
                    .collect();
                // ⛔ THE ELEMENT TYPES MUST AGREE TOO, and this arm only ever
                // compared ADDRESS WIDTHS — so `table.copy $funcs $externs`
                // typed clean. `table_of` returned the element type all along;
                // nothing read it.
                let (dst, src) = if ids.len() >= 2 {
                    let (_, de, da) = self.table_of(&ids[0..1])?;
                    let (_, se, sa) = self.table_of(&ids[1..2])?;
                    if !vt_subtype(&se, &de, &self.ctx.types) {
                        return tmis_d(format!(
                            "table element {} is not a subtype of {}",
                            vt_show(&se),
                            vt_show(&de)
                        ));
                    }
                    (da, sa)
                } else {
                    let a = self.table_of(&[])?.2;
                    (a.clone(), a)
                };
                let len = if matches!(dst, Vt::Num("i32")) || matches!(src, Vt::Num("i32")) {
                    Vt::Num("i32")
                } else {
                    dst.clone()
                };
                self.pop_expect(&len)?;
                self.pop_expect(&src)?;
                self.pop_expect(&dst)?;
                Ok(())
            }
            "table.init" => {
                // `table.init $t $el` names a TABLE then an ELEMENT SEGMENT,
                // and the segment's type must be a subtype of the table's
                // element type. With one immediate the table is 0.
                let ids: Vec<String> = imms
                    .iter()
                    .filter(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                    .cloned()
                    .collect();
                let (tbl, seg) = if ids.len() >= 2 {
                    (&ids[0..1], Some(&ids[1]))
                } else {
                    (&ids[0..0], ids.first())
                };
                let (_, elem, a) = self.table_of(tbl)?;
                if let Some(si) = seg.and_then(|k| resolve_wast_index(k, &self.ctx.elem_names)) {
                    if let Some(st) = self.ctx.elem_types.get(si).cloned().flatten() {
                        if !vt_subtype(&st, &elem, &self.ctx.types) {
                            return tmis_d(format!(
                                "element segment {} is not a subtype of table element {}",
                                vt_show(&st),
                                vt_show(&elem)
                            ));
                        }
                    }
                }
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&a)?;
                Ok(())
            }

            // ── GC ───────────────────────────────────────────────────────
            //
            // ⛔ EVERY ARM RESOLVES ITS TYPE OR BAILS. A GC instruction's
            // operand and result types come from the TYPE TABLE, not the
            // mnemonic, so an unresolvable index or an unparseable field
            // spelling abandons the function rather than guessing — the same
            // discipline the rest of the pass runs on.
            "struct.new" | "struct.new_default" => {
                let ti = self.type_index(imms.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                if t.kind != Some("struct") {
                    return tmis_d(format!("type {ti} is not a struct"));
                }
                if name == "struct.new" {
                    // Fields are pushed in declaration order, so they pop in
                    // reverse.
                    let fields = self.field_types(&t)?;
                    self.pop_vals(&fields)?;
                }
                self.push_val(Vt::Ref(RefT {
                    nullable: false,
                    exact: true,
                    heap: Heap::Concrete(ti),
                }));
                Ok(())
            }
            "struct.get" | "struct.get_s" | "struct.get_u" | "struct.set" => {
                let ti = self.type_index(imms.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                if t.kind != Some("struct") {
                    return tmis_d(format!("type {ti} is not a struct"));
                }
                // The field is named by the SECOND immediate, in EITHER
                // spelling.
                //
                // ⛔ THE NAMED ONE USED TO BAIL, AND A BAIL IS INVISIBLE. Only
                // `parse::<usize>()` was tried, so `struct.get 0 $x` abandoned
                // the whole function instead of typing it — the third time in
                // this file that one spelling carried a rule its sibling did
                // not (`global.set $g` vs `global.set 0`, `throw $t` vs
                // `throw 0`). Resolution is per-TYPE, so it reads this type's
                // own `field_names`.
                let key = imms.get(1).ok_or(Fail::Bail)?;
                let fi: usize = match key.parse::<usize>() {
                    Ok(n) => n,
                    Err(_) => t
                        .field_names
                        .iter()
                        .position(|n| n.as_deref() == Some(key.as_str()))
                        .ok_or(Fail::Bail)?,
                };
                let spelling = t.field_types.get(fi).cloned().ok_or(Fail::Bail)?;
                let packed = matches!(strip_mut(&spelling).as_str(), "i8" | "i16");
                // ⛔ A PACKED FIELD HAS NO PLAIN ACCESSOR. `i8`/`i16` are
                // storage-only: they must be read through `_s`/`_u`, which
                // yield i32, and a bare `struct.get` on one is invalid.
                if packed && name == "struct.get" {
                    return tmis_d("packed field needs struct.get_s or struct.get_u".to_string());
                }
                if !packed && matches!(name, "struct.get_s" | "struct.get_u") {
                    return tmis_d("sign extension on a non-packed field".to_string());
                }
                let fty = if packed {
                    Vt::Num("i32")
                } else {
                    parse_vt(&strip_mut(&spelling), &self.ctx.type_names).ok_or(Fail::Bail)?
                };
                let r = Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(ti),
                });
                if name == "struct.set" {
                    // ⛔ MUTABILITY IS OPT-IN PER FIELD, exactly as it is for an
                    // array element. `(field i32)` is NOT `(field (mut i32))`,
                    // so writing it is invalid however well-typed the value is
                    // — a property of the TYPE, with its own diagnostic.
                    if !spelling.trim_start().starts_with("(mut") {
                        return Err(Fail::Mismatch("immutable field".to_string()));
                    }
                    self.pop_expect(&fty)?;
                    self.pop_expect(&r)?;
                } else {
                    self.pop_expect(&r)?;
                    self.push_val(fty);
                }
                Ok(())
            }
            "array.new" | "array.new_default" | "array.new_fixed" => {
                let ti = self.type_index(imms.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                if t.kind != Some("array") {
                    return tmis_d(format!("type {ti} is not an array"));
                }
                let elem = self.array_elem_vt(&t)?;
                match name {
                    "array.new" => {
                        self.pop_expect(&Vt::Num("i32"))?;
                        self.pop_expect(&elem)?;
                    }
                    "array.new_default" => {
                        self.pop_expect(&Vt::Num("i32"))?;
                    }
                    _ => {
                        // `array.new_fixed $t N` — N elements, all on the stack.
                        let n: usize =
                            imms.get(1).and_then(|x| x.parse().ok()).ok_or(Fail::Bail)?;
                        for _ in 0..n {
                            self.pop_expect(&elem)?;
                        }
                    }
                }
                self.push_val(Vt::Ref(RefT {
                    nullable: false,
                    exact: true,
                    heap: Heap::Concrete(ti),
                }));
                Ok(())
            }
            "array.get" | "array.get_s" | "array.get_u" | "array.set" => {
                let ti = self.type_index(imms.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                if t.kind != Some("array") {
                    return tmis_d(format!("type {ti} is not an array"));
                }
                let raw = t.array_elem.clone().ok_or(Fail::Bail)?;
                let packed = matches!(raw.as_str(), "i8" | "i16");
                if packed && name == "array.get" {
                    return tmis_d("packed element needs array.get_s or array.get_u".to_string());
                }
                if !packed && matches!(name, "array.get_s" | "array.get_u") {
                    return tmis_d("sign extension on a non-packed element".to_string());
                }
                let elem = self.array_elem_vt(&t)?;
                let r = Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(ti),
                });
                if name == "array.set" {
                    self.pop_expect(&elem)?;
                    self.pop_expect(&Vt::Num("i32"))?;
                    self.pop_expect(&r)?;
                } else {
                    self.pop_expect(&Vt::Num("i32"))?;
                    self.pop_expect(&r)?;
                    self.push_val(elem);
                }
                Ok(())
            }
            "array.len" => {
                let t = self.pop_val()?;
                if !Self::is_ref(&t) {
                    return tmis_d("array.len needs a reference".to_string());
                }
                self.push_val(Vt::Num("i32"));
                Ok(())
            }
            "array.fill" => {
                let ti = self.type_index(imms.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                let elem = self.array_elem_vt(&t)?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&elem)?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(ti),
                }))?;
                Ok(())
            }
            // §3.4.8.14: the SOURCE array's storage type must be a subtype of
            // the DESTINATION's.
            //
            // ⛔ PACKEDNESS IS PART OF THE STORAGE TYPE, AND `array_elem_vt`
            // ERASES IT — it reads i8 and i16 as the i32 they load as, which is
            // right for a value and wrong here. Both fixtures are invalid only
            // in the erased detail: `(mut i8)` from `i16` agrees once both are
            // i32, and `(mut i8)` from `(mut (ref $a))` needs the packed side
            // to reject a reference outright. So this compares the RAW
            // spellings, and a packed type matches only itself.
            "array.copy" => {
                let ids: Vec<String> = imms
                    .iter()
                    .filter(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                    .cloned()
                    .collect();
                if ids.len() < 2 {
                    return Err(Fail::Bail);
                }
                let di = self.type_index(Some(&ids[0]))?;
                let si = self.type_index(Some(&ids[1]))?;
                let dt = self.ctx.types.get(di).cloned().ok_or(Fail::Bail)?;
                let st = self.ctx.types.get(si).cloned().ok_or(Fail::Bail)?;
                if dt.kind != Some("array") || st.kind != Some("array") {
                    return tmis_d("array.copy needs two array types".to_string());
                }
                if !dt.array_elem_mut {
                    return Err(Fail::Mismatch("array is immutable".to_string()));
                }
                let draw = dt.array_elem.clone().ok_or(Fail::Bail)?;
                let sraw = st.array_elem.clone().ok_or(Fail::Bail)?;
                let dpk = matches!(draw.as_str(), "i8" | "i16");
                let spk = matches!(sraw.as_str(), "i8" | "i16");
                let agree = if dpk || spk {
                    draw == sraw
                } else {
                    let dv = self.array_elem_vt(&dt)?;
                    let sv = self.array_elem_vt(&st)?;
                    vt_subtype(&sv, &dv, &self.ctx.types)
                };
                if !agree {
                    return Err(Fail::Mismatch("array types do not match".to_string()));
                }
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(si),
                }))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(di),
                }))?;
                Ok(())
            }
            // §3.4.8.16-17: `array.init_data` fills from raw bytes, so the
            // element must be numeric or vector — a reference has no byte
            // spelling. `array.init_elem` fills from an element segment, so the
            // segment's type must be a subtype of the element type.
            "array.init_data" | "array.init_elem" => {
                let ids: Vec<String> = imms
                    .iter()
                    .filter(|t| t.starts_with('$') || t.parse::<usize>().is_ok())
                    .cloned()
                    .collect();
                let ti = self.type_index(ids.first())?;
                let t = self.ctx.types.get(ti).cloned().ok_or(Fail::Bail)?;
                if t.kind != Some("array") {
                    return tmis_d(format!("type {ti} is not an array"));
                }
                if !t.array_elem_mut {
                    return Err(Fail::Mismatch("array is immutable".to_string()));
                }
                let elem = self.array_elem_vt(&t)?;
                if name == "array.init_data" {
                    if Self::is_ref(&elem) {
                        return Err(Fail::Mismatch(
                            "array type is not numeric or vector".to_string(),
                        ));
                    }
                } else {
                    // ⛔ RESOLVED THROUGH THE SEGMENT INDEX SPACE, NOT THE TYPE
                    // ONE. `$e1` names an element segment here; looking it up
                    // among types would find nothing and abstain, which is how
                    // both fixtures went undetected.
                    let si = ids
                        .get(1)
                        .and_then(|k| resolve_wast_index(k, &self.ctx.elem_names))
                        .ok_or(Fail::Bail)?;
                    let seg = self
                        .ctx
                        .elem_types
                        .get(si)
                        .cloned()
                        .flatten()
                        .ok_or(Fail::Bail)?;
                    if !vt_subtype(&seg, &elem, &self.ctx.types) {
                        return Err(Fail::Mismatch(format!(
                            "type mismatch: instruction requires [{}] but stack has [{}]",
                            vt_show(&elem),
                            vt_show(&seg)
                        )));
                    }
                }
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Num("i32"))?;
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Concrete(ti),
                }))?;
                Ok(())
            }
            "ref.eq" => {
                let eq = Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("eq"),
                });
                self.pop_expect(&eq)?;
                self.pop_expect(&eq)?;
                self.push_val(Vt::Num("i32"));
                Ok(())
            }
            "ref.i31" => {
                self.pop_expect(&Vt::Num("i32"))?;
                self.push_val(Vt::Ref(RefT {
                    nullable: false,
                    exact: false,
                    heap: Heap::Abs("i31"),
                }));
                Ok(())
            }
            "i31.get_s" | "i31.get_u" => {
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("i31"),
                }))?;
                self.push_val(Vt::Num("i32"));
                Ok(())
            }
            "any.convert_extern" => {
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("extern"),
                }))?;
                self.push_val(Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("any"),
                }));
                Ok(())
            }
            "extern.convert_any" => {
                self.pop_expect(&Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("any"),
                }))?;
                self.push_val(Vt::Ref(RefT {
                    nullable: true,
                    exact: false,
                    heap: Heap::Abs("extern"),
                }));
                Ok(())
            }

            _ => {
                // v128 lane-indexed accesses carry the vector as an operand.
                if let Some(rest) = name.strip_prefix("v128.load") {
                    if rest.ends_with("_lane") {
                        let a = self.mem_addr(imms)?;
                        self.pop_expect(&Vt::Num("v128"))?;
                        self.pop_expect(&a)?;
                        self.push_val(Vt::Num("v128"));
                        return Ok(());
                    }
                }
                if let Some(rest) = name.strip_prefix("v128.store") {
                    if rest.ends_with("_lane") {
                        let a = self.mem_addr(imms)?;
                        self.pop_expect(&Vt::Num("v128"))?;
                        self.pop_expect(&a)?;
                        return Ok(());
                    }
                }
                if let Some((vt, is_store)) = mem_access_sig(name) {
                    let a = self.mem_addr(imms)?;
                    if is_store {
                        self.pop_expect(&Self::num(vt))?;
                        self.pop_expect(&a)?;
                    } else {
                        self.pop_expect(&a)?;
                        self.push_val(Self::num(vt));
                    }
                    return Ok(());
                }
                if let Some((p, r)) = fixed_sig(name) {
                    let ps: Vec<Vt> = p.iter().map(|t| Self::num(t)).collect();
                    self.pop_vals(&ps)?;
                    for t in r {
                        self.push_val(Self::num(t));
                    }
                    return Ok(());
                }
                // ⛔ NEVER GUESS. An unmodelled mnemonic abandons the whole
                // function — GC, exceptions, atomics and the string builtins
                // all land here on purpose.
                Err(Fail::Bail)
            }
        }
    }
}

// ── The driver ───────────────────────────────────────────────────────────────

/// Type a module's INITIALIZER const-expressions against their declared type.
///
/// ⛔ AN INITIALIZER IS TYPED TOO, AND NOTHING WAS TYPING IT. A table's inline
/// init and a global's init are constant EXPRESSIONS whose result must match
/// the declared element/global type — `(table 1 (ref null func) (i32.const 0))`
/// is invalid because an i32 is not a funcref. The stack-typing pass walks
/// FUNCTION BODIES only, so these went unexamined: 10 fixtures in table.wast
/// alone.
///
/// Reuses the same machinery — an initializer is just a body with no locals
/// whose declared result is the entity's type — so it bails identically on
/// anything it cannot type.
fn module_init_expr_reason(module: &Pair<Rule>) -> Option<String> {
    let ctx = build_type_ctx(module);
    // How far into the global index space the NEXT defined global sits.
    let mut global_at = ctx.imported_globals;
    for field in module_fields(module) {
        // An IMPORTED entity has no initializer.
        if find_rule(&field, Rule::import_inline).is_some() {
            continue;
        }
        // ⛔ HOW MANY GLOBALS A CONSTANT EXPRESSION MAY READ DEPENDS ON WHERE
        // IT SITS, and the reason is INSTANTIATION ORDER (§4.5.4), not a
        // uniform rule about constant expressions:
        //
        //   * a TABLE is allocated BEFORE any global is initialised, so its
        //     initializer sees the IMPORTED globals only — the suite asserts
        //     `(global $g funcref …) (table $t 10 funcref (global.get $g))`
        //     is "unknown global" even though `$g` is declared first;
        //   * a GLOBAL at index i sees `[0, i)` — imports and the globals
        //     already initialised, which is why a forward reference is
        //     "unknown global" but a backward one is fine;
        //   * an ELEM or DATA offset runs after every global exists, so it
        //     sees all of them.
        //
        // Reading this as one rule ("imports only") rejected SEVEN valid
        // modules across global/elem/data — invisible to the suite, since
        // validation runs only inside `assert_invalid`, and caught only by
        // wrapping the suite's VALID modules in one.
        let limit = match field.as_rule() {
            Rule::table_field => ctx.imported_globals,
            Rule::global_field => global_at,
            _ => ctx.globals.len(),
        };
        if field.as_rule() == Rule::global_field {
            global_at += 1;
        }
        let want = match field.as_rule() {
            Rule::table_field => {
                // The element type is the `ref_val_type` inside `table_type`.
                let tt = find_rule(&field, Rule::table_type)?;
                tt.into_inner()
                    .filter(|c| !matches!(c.as_rule(), Rule::integer))
                    .last()
                    .and_then(|c| parse_vt(c.as_str().trim(), &ctx.type_names))
            }
            Rule::global_field => find_rule(&field, Rule::global_type).and_then(|g| {
                let t = g.as_str().trim();
                let spelling = t
                    .strip_prefix('(')
                    .and_then(|x| x.trim().strip_prefix("mut"))
                    .map(|x| x.trim().trim_end_matches(')').trim())
                    .unwrap_or(t);
                parse_vt(spelling, &ctx.type_names)
            }),
            // ⛔ A SEGMENT OFFSET IS A CONST-EXPRESSION TOO, and its type is
            // the TARGET's address width — `(table 1 funcref) (elem
            // (i64.const 0))` is invalid because that table is i32-indexed.
            // Only an ACTIVE segment has one.
            Rule::elem_field => ctx.table_addr.first().map(|t| Vt::Num(t)),
            Rule::data_field => ctx.mem_addr.first().map(|t| Vt::Num(t)),
            _ => continue,
        };
        let Some(want) = want else { continue };
        let init: Vec<Pair<Rule>> = if matches!(field.as_rule(), Rule::elem_field | Rule::data_field)
        {
            // The offset lives inside `elem_mode`/`data_mode`. ⛔ AND THE SAME
            // GRAMMAR AMBIGUITY BITES HERE: `elem_mode`'s last alternative is a
            // bare folded instruction, so `(elem $e (ref 1))`'s element TYPE
            // parses as an offset. `elem_mode_is_reference_type` is what tells
            // them apart — without it every typed passive segment reports a
            // mismatch against the table's address type.
            let mode = field
                .clone()
                .into_inner()
                .find(|c| matches!(c.as_rule(), Rule::elem_mode | Rule::data_mode));
            match mode {
                Some(m)
                    if !(m.as_rule() == Rule::elem_mode && elem_mode_is_reference_type(&m))
                        && m.as_str().trim() != "declare" =>
                {
                    m.into_inner()
                        .filter(|c| matches!(c.as_rule(), Rule::folded_instr | Rule::instr))
                        .collect()
                }
                _ => Vec::new(),
            }
        } else {
            field
                .clone()
                .into_inner()
                .filter(|c| matches!(c.as_rule(), Rule::folded_instr | Rule::instr))
                .collect()
        };
        if init.is_empty() {
            // ⛔ EMPTY IS ONLY LEGAL WHERE THERE IS NOTHING TO INITIALISE.
            // A global's initializer is REQUIRED — `(global i32)` delivers no
            // value where an i32 is wanted, which the suite asserts as a type
            // mismatch. A table has no initializer at all (`(table 1 funcref)`
            // is perfectly valid), and a PASSIVE segment has no offset, so
            // those two must keep skipping or every valid module trips here.
            if field.as_rule() == Rule::global_field {
                return Some("type mismatch".to_string());
            }
            continue;
        }
        let mut flat = Vec::new();
        if flatten_instrs(init, &mut flat).is_none() {
            continue;
        }
        let mut tv = Tv {
            ctx: &ctx,
            vals: Vec::new(),
            ctrls: Vec::new(),
            locals: Vec::new(),
            local_names: HashMap::new(),
            inited: Vec::new(),
            global_limit: Some(limit),
        };
        let results = vec![want];
        tv.push_ctrl("func", None, Vec::new(), results.clone());
        let mut bailed = false;
        for ins in &flat {
            match tv.step(ins, &results) {
                Ok(()) => {}
                Err(Fail::Mismatch(m)) => return Some(m),
                Err(Fail::Bail) => {
                    bailed = true;
                    break;
                }
            }
        }
        if bailed || tv.ctrls.len() != 1 {
            continue;
        }
        if let Err(Fail::Mismatch(m)) = tv.pop_ctrl() {
            return Some(m);
        }
    }
    None
}

/// §3.4.7: the start function takes no parameters and returns no results.
/// `(func $main (result i32) …) (start $main)` and `(func $main (param i32))`
/// are both invalid, and the suite gives that its own wording — "start
/// function", not "type mismatch".
///
/// Separate from the EXISTENCE check in `name_resolution_walk`: that one runs
/// off the census, which counts declarations without their signatures. This
/// needs the typed index space, so it lives where the `TypeCtx` is built.
fn start_function_type_reason(module: &Pair<Rule>) -> Option<String> {
    let ctx = build_type_ctx(module);
    for field in module_fields(module) {
        let inner = match field.as_rule() {
            Rule::module_field => field.clone().into_inner().next()?,
            _ => field.clone(),
        };
        if inner.as_rule() != Rule::start_field {
            continue;
        }
        let idx = inner
            .clone()
            .into_inner()
            .find(|c| c.as_rule() == Rule::index)?;
        let i = resolve_wast_index(idx.as_str().trim(), &ctx.func_names)?;
        // An unresolvable or untypeable function abstains — "unknown function"
        // is a different rule's answer, and guessing here would give the wrong
        // wording to a module that is invalid for another reason entirely.
        let sig = ctx.funcs.get(i)?.as_ref()?;
        if !sig.params.is_empty() || !sig.results.is_empty() {
            return Some("start function".to_string());
        }
    }
    None
}

/// Type every defined function in a module. `None` means nothing was proved —
/// either the module types clean or some function could not be typed at all.
fn module_stack_typing_reason(module: &Pair<Rule>) -> Option<String> {
    let ctx = build_type_ctx(module);
    for field in module_fields(module) {
        if field.as_rule() != Rule::func_field {
            continue;
        }
        // An imported function has no body to type.
        if find_rule(&field, Rule::import_inline).is_some() {
            continue;
        }
        if let Some(m) = type_one_func(&field, &ctx) {
            return Some(m);
        }
    }
    None
}

/// A function's locals, params first — the one index space `local.get` reads.
///
/// ⛔ A `(type $t)`-ONLY FUNCTION STILL HAS PARAMETERS. `func_locals` reads the
/// DECLARED `param`/`local` groups, so a body written against a type index has
/// none of them and every `local.get` shifts by the parameter count — silently
/// typing `local.get 0` as the first LOCAL. Prepend the signature's params and
/// shift the name map with them.
fn func_local_types(
    f: &Pair<Rule>,
    tu: &Pair<Rule>,
    sig: &FnSig,
    names: &HashMap<String, usize>,
) -> Option<(Vec<Vt>, HashMap<String, usize>, Vec<bool>)> {
    let has_inline_params = tu.clone().into_inner().any(|c| c.as_rule() == Rule::param);
    let (declared, declared_names, declared_inited) = func_locals(f, names)?;
    if has_inline_params {
        return Some((declared, declared_names, declared_inited));
    }
    let off = sig.params.len();
    let mut v = sig.params.clone();
    v.extend(declared);
    let shifted = declared_names.into_iter().map(|(k, i)| (k, i + off)).collect();
    // The signature's params come first and are initialized by arrival.
    let mut inited = vec![true; off];
    inited.extend(declared_inited);
    Some((v, shifted, inited))
}

/// A NUMERIC `local.get`/`set`/`tee` index past the end of the function's
/// locals.
///
/// ⛔ `name_resolution_walk` DECLINES THIS CASE ON PURPOSE — its own comment
/// says a numeric local index "needs the param+local count, which the typeuse
/// may state by reference to a type". That count exists now: the type context
/// resolves a `typeuse` through both spellings AND the implicit types, so the
/// check is answerable where it was not before. 19 fixtures assert it.
///
/// Reports NOTHING when the count cannot be established — the same bail as
/// everywhere else, because inventing a bound would reject valid modules.
fn module_unknown_local_reason(module: &Pair<Rule>) -> Option<String> {
    let ctx = build_type_ctx(module);
    for f in module_fields(module) {
        if f.as_rule() != Rule::func_field || find_rule(&f, Rule::import_inline).is_some() {
            continue;
        }
        let Some(tu) = find_rule(&f, Rule::typeuse) else { continue };
        let Some((_, sig)) = sig_from_typeuse(&tu, &ctx.type_sigs, &ctx.type_names, &ctx.type_names)
        else {
            continue;
        };
        let Some((locals, _, _)) = func_local_types(&f, &tu, &sig, &ctx.type_names) else {
            continue;
        };
        if let Some(n) = first_out_of_range_local(&f, locals.len()) {
            return Some(format!("unknown local {n}"));
        }
    }
    None
}

fn first_out_of_range_local(p: &Pair<Rule>, count: usize) -> Option<u128> {
    if matches!(p.as_rule(), Rule::plain_instr | Rule::folded_instr) {
        let head = instr_head_name(p).unwrap_or_default();
        if matches!(head.as_str(), "local.get" | "local.set" | "local.tee") {
            for c in p.clone().into_inner() {
                if c.as_rule() != Rule::instr_arg {
                    continue;
                }
                // A `$name` is another rule's business; only a written NUMBER
                // is decidable from a count.
                let t = c.as_str().trim();
                if t.starts_with('$') {
                    continue;
                }
                if let Some(v) = parse_wat_u128(t) {
                    if v >= count as u128 {
                        return Some(v);
                    }
                }
            }
        }
    }
    for c in p.clone().into_inner() {
        if let Some(v) = first_out_of_range_local(&c, count) {
            return Some(v);
        }
    }
    None
}

fn type_one_func(f: &Pair<Rule>, ctx: &TypeCtx) -> Option<String> {
    let names = &ctx.type_names;
    let tu = find_rule(f, Rule::typeuse)?;
    let (_, sig) = sig_from_typeuse(&tu, &ctx.type_sigs, &ctx.type_names, names)?;

    // ⛔ A `(type $t)`-ONLY FUNCTION STILL HAS PARAMETERS. `func_locals` reads
    // the DECLARED param/local groups, so a body written against a type index
    // has none of them and every `local.get` shifts by the parameter count —
    // silently typing `local.get 0` as the first local. Prepend the signature's
    // params and shift the name map with them.
    let (locals, local_names, inited) = func_local_types(f, &tu, &sig, names)?;

    let body: Vec<Pair<Rule>> = f
        .clone()
        .into_inner()
        .filter(|c| c.as_rule() == Rule::instr)
        .collect();
    let mut flat = Vec::new();
    flatten_instrs(body, &mut flat)?;

    let mut tv = Tv {
        ctx,
        vals: Vec::new(),
        ctrls: Vec::new(),
        locals,
        local_names,
        inited,
        // A function body reads the whole global index space.
        global_limit: None,
    };
    // The function body is itself a block returning the function's results.
    tv.push_ctrl("func", None, Vec::new(), sig.results.clone());
    for ins in &flat {
        match tv.step(ins, &sig.results) {
            Ok(()) => {}
            Err(Fail::Mismatch(m)) => return Some(m),
            Err(Fail::Bail) => return None,
        }
    }
    // An unbalanced body is MALFORMED, not invalid — a different command's
    // question, so it leaves here reporting nothing.
    if tv.ctrls.len() != 1 {
        return None;
    }
    match tv.pop_ctrl() {
        Ok(_) => None,
        Err(Fail::Mismatch(m)) => Some(m),
        Err(Fail::Bail) => None,
    }
}

/// Find the modules under a pair and type each. Used from the two places an
/// `assert_invalid` module arrives — inline and `(module quote …)`.
fn stack_typing_reason_in(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() == Rule::module {
        // Initializers first: they are typed independently of any function,
        // and a module whose table init is ill-typed says so regardless of
        // what its bodies do.
        if let Some(r) = module_init_expr_reason(pair) {
            return Some(r);
        }
        if let Some(r) = start_function_type_reason(pair) {
            return Some(r);
        }
        return module_stack_typing_reason(pair);
    }
    for c in pair.clone().into_inner() {
        if let Some(r) = stack_typing_reason_in(&c) {
            return Some(r);
        }
    }
    None
}
