//! The module's global namespace as ONE object — and the logistics that go
//! with it: which names the module OWNS, and which it merely reads.
//!
//! Four languages spell the same thing — Lua `_G`, JS `globalThis`, PHP
//! `$GLOBALS`, Python `globals()` (the one spelled as a call). Those words
//! appear NOWHERE below: each walker normalizes its own into
//! `ExprKind::GlobalNamespace` and this module only ever sees the node. They
//! are named here in prose, once, so a reader knows what the node stands for —
//! not so anything can match on them.
//!
//! All four name the module's OBJECT environment record (ECMA-262 §9.1.1.4:
//! a Global Environment Record is an object record plus a declarative one —
//! `var` and function declarations land in the object half, `let`/`const` in
//! the declarative half, which is why `globalThis.x` sees the first and not
//! the second). It is the same storage a bare `x` reads, so a namespace object
//! is a *view* over storage that already exists; nothing new has to be stored.
//!
//! ⛔ **This paragraph used to describe a model that no longer exists**, and
//! said so as fact: *"`vm.globals` is already a live `HashMap<String, Value>`
//! … `GLOBAL_GET` is nothing but resolve a constant index to a string, look it
//! up in that map."* That was true before the globalidx migration and is not
//! now — `vm.globals` is a `Vec<Value>`, `GLOBAL_GET` reads a real globalidx
//! over `global_imports ++ defined` and its own comment says *"No name is
//! consulted here"*, and `global_index: HashMap<String, u32>` survives only to
//! bind host imports AT INSTANTIATION.
//!
//! Left as a marker because `directives.md` §1 names this exact cost: a
//! document that describes a retired mechanism as live gets cited by the next
//! reader, who then repeats a claim the tree has already disproved. When a
//! mechanism is replaced, the comment naming it is part of the replacement.
//!
//! Before this module each language had built its own compile-time workaround
//! and none could share: JS rewrote `globalThis.X` in `expressions.rs`, Lua
//! rewrote `_G.x` in its walker, Python materialised a SNAPSHOT dict from
//! `defined_globals`, and PHP resolved `$$name` against `__php_var_vars` — a
//! parallel table that the real globals never see. Four answers to one
//! question, three of them wrong in a different way (measured 2026-08-02: Lua
//! `_G[k]` crashed, Python's write never took, PHP's computed read came back
//! blank).
//!
//! ## What this module does and does not cover
//!
//! The LITERAL-key path is here: `_G.x` / `globalThis.x` / `$GLOBALS['x']`
//! compile to a direct `GLOBAL_GET`/`GLOBAL_SET`, which is what all four were
//! already doing by hand. That is a pure de-duplication — no behaviour changes.
//!
//! A COMPUTED key (`_G[k]`) and enumeration (`for k in globals()`) are NOT
//! here. `GLOBAL_GET` takes the name as a u16 immediate, so a runtime key
//! cannot reach the map through it, and the host surface is closed — no host
//! function may be added to reflect `vm.globals`. Closing that gap means
//! backing the namespace with a real dict (`primitives/dict.rs`, which already
//! carries `__keys` insertion order) and compiling member bindings into it, so
//! the object IS the storage rather than a view of it. That is a codegen change
//! for every global access in these languages and is deliberately not bundled
//! in here.

use super::*;
use vybe_runtime::profile::LanguageProfile;

/// Where the module's one namespace object lives.
///
/// A hidden module global rather than a local, because the object has to be
/// the SAME one at every site that names it — including sites in different
/// chunks, which a local cannot reach. `declare_free_globals` declares it like
/// any other free name, and it is invisible to
/// `Compiler::global_namespace_members`, which reads declared bindings and
/// module scope, so the namespace never contains itself.
const GLOBAL_NAMESPACE_CACHE: &str = "__vybe_global_namespace";

/// How a profile spells its global namespace, and which bindings belong to it.
pub struct Options {
    /// Bindings introduced by `let`/`const`/`class` live in the DECLARATIVE
    /// record and are NOT members of the namespace object — ECMA-262 §16.1.7
    /// puts only `var` and function declarations on the global object. Lua,
    /// Python and PHP have no such split: every module-level name is a member.
    ///
    /// Verified against `node` and vybe on 2026-08-02, which already agree:
    /// `var a=1; let b=1; const c=1; function f(){}` →
    /// `globalThis.a` is `1`, `b`/`c` are `undefined`, `f` is a function.
    pub lexical_bindings_are_members: bool,
}

impl Options {
    pub fn from_profile(profile: &LanguageProfile) -> Options {
        Options {
            lexical_bindings_are_members: !profile.has_ecma_globals,
        }
    }
}

impl Compiler {
    /// Every module-level name this namespace object exposes, in a stable
    /// order. `defined_globals` is the link pass's record of the source's own
    /// top-level declarations, which is exactly the membership question.
    fn global_namespace_members(&self) -> Vec<String> {
        let mut names: Vec<String> = self.defined_globals.iter().cloned().collect();
        names.extend(self.module_variable_names.iter().cloned());
        // A module-level binding is a member whether the backend stored it as a
        // VM global or as a local of the module chunk. PHP's top-level `$x = 5`
        // is the latter — it never enters `defined_globals`, so a members list
        // built from that alone made every `$GLOBALS[$k]` lookup miss and read
        // blank. `scopes[0]` is the module scope; deeper scopes are function
        // locals and are NOT members.
        if let Some(module_scope) = self.scopes.first() {
            names.extend(
                module_scope
                    .defined_names
                    .iter()
                    // Compiler temporaries (`__php_offset_obj_1`, `__tmp`) are
                    // not user bindings and must not be exposed. This asks the
                    // BINDING who declared it — recorded by
                    // `define_source_local` where the fact is known — rather
                    // than inferring provenance from the spelling. A user
                    // variable named `__x` is a member; a temporary that
                    // forgot the prefix still is not.
                    .filter(|entry| entry.is_source())
                    .map(|entry| entry.name.clone()),
            );
        }
        names.sort();
        names.dedup();
        // A language that spells variables in their own namespace exposes only
        // VARIABLES here: real PHP's `$GLOBALS` holds `$x` but not the function
        // `foo`, which vybe was returning. Languages with no marker (Python,
        // Lua) put functions and classes in the namespace too, so they keep
        // every name.
        if self.variable_namespace.is_some() {
            names.retain(|name| self.is_variable_name(name));
        }
        names
    }

    /// `ns[key]` where `key` is only known at runtime.
    ///
    /// `GLOBAL_GET` takes the name as a u16 IMMEDIATE, so a runtime key cannot
    /// index the map directly, and the host surface is closed — no host fn may
    /// be added to reflect `vm.globals`. What composes from existing ops is a
    /// comparison per known member, each guarding a real `GLOBAL_GET`. That
    /// keeps the read LIVE (unlike a snapshot, it sees the current value) at
    /// the cost of O(members) comparisons — paid only where a computed-key
    /// access actually appears in the source.
    ///
    /// A name the module never declared reads as undefined, matching a miss.
    pub(super) fn emit_global_namespace_index_get(
        &mut self,
        key: &Expression,
    ) -> Result<(), String> {
        let line = self.line;
        self.compile_expr(key)?;
        let key_slot = self.define_local("__globals_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        let result_slot = self.define_local("__globals_result");
        crate::primitives::instructions::core_wasm::undefined(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        for name in self.global_namespace_members() {
            // The KEY is spelled as source text (`"x"`), while a member of a
            // language with a separate variable namespace is stored marked
            // (`$x` in PHP). Compare on the body, read with the full name —
            // otherwise every PHP lookup misses and reads undefined.
            let key_text = self.variable_name_body(&name).to_string();
            self.emit_u16(Op::LOCAL_GET, key_slot);
            self.emit_const(Value::String(Arc::from(key_text.as_str())));
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_var_get(&name);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    /// `ns[key] = value`, the write twin of
    /// [`Self::emit_global_namespace_index_get`]. The value is already on the
    /// stack.
    ///
    /// Limit worth knowing: this can only assign to a name the module
    /// DECLARED. Real Python/Lua also let you create a brand-new global through
    /// the namespace (`globals()['fresh'] = 1`), which needs a runtime-keyed
    /// store the closed host surface cannot express.
    pub(super) fn emit_global_namespace_index_set(
        &mut self,
        key: &Expression,
    ) -> Result<(), String> {
        let line = self.line;
        let value_slot = self.define_local("__globals_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.compile_expr(key)?;
        let key_slot = self.define_local("__globals_set_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        for name in self.global_namespace_members() {
            let key_text = self.variable_name_body(&name).to_string();
            self.emit_u16(Op::LOCAL_GET, key_slot);
            self.emit_const(Value::String(Arc::from(key_text.as_str())));
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_var_set(&name);
            self.chunk().emit_end(line);
        }
        Ok(())
    }

    /// The namespace itself as an OBJECT — `globals()` passed around,
    /// `for k in globals()`, `pairs(_G)`, `foreach ($GLOBALS as $k => $v)`.
    ///
    /// Built ONCE per module and cached in [`GLOBAL_NAMESPACE_CACHE`], because
    /// the namespace is a singleton and every language here says so: lua's
    /// `_G == _G` is true, python's `globals() is globals()` is true, and
    /// ECMA-262 §19.3.1 requires one global object per realm. Rebuilding a
    /// fresh dict per evaluation is not a slower way to be right — it is
    /// wrong, and `lua/globals/g_is_same_reference_as_global_env` is the test
    /// that says so.
    ///
    /// Values are re-read on EVERY evaluation, so the object tracks writes the
    /// way a member read does — including through an alias taken earlier
    /// (measured: `g1 = globals(); zz = 2` leaves `zz` visible in `g1`). WHERE
    /// a name is declared does not matter either, because membership comes
    /// from `global_namespace_members`, which the link pass has already
    /// finished; only the values are runtime.
    ///
    /// The honest limit is that membership is what the module DECLARES. A name
    /// that exists only at runtime is not a member — and cannot be made one,
    /// since `emit_global_namespace_index_set` can only assign to a declared
    /// name for the same reason. Closing both means backing the namespace with
    /// the dict itself so the object IS the storage — the codegen change this
    /// module's header already scopes out.
    ///
    /// ⛔ This was python's alone, keyed by the NAME `globals` in
    /// `builtins.rs::try_compile_builtin` — a language spelling sitting in the
    /// shared builtin table, and its own comment admitted the hazard: *"A user
    /// function named `globals()` in another language would now reach here."*
    /// Lua's `_G` and php's `$GLOBALS` had no object at all, so `pairs(_G)` and
    /// `foreach ($GLOBALS ...)` enumerated NOTHING (measured 2026-08-27: 0
    /// against real lua's and php's 2). One node, one emit, four languages.
    pub(super) fn emit_global_namespace_object(&mut self) {
        let line = self.line;
        let members = self.global_namespace_members();

        // Create on first evaluation only. A never-written global reads as
        // undefined, which `emit_dyn_not` turns into the i32 the `if` wants.
        self.emit_global_read(GLOBAL_NAMESPACE_CACHE);
        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        self.emit_global_write(GLOBAL_NAMESPACE_CACHE);
        for name in &members {
            // Same split as the computed-key path: the KEY is source text
            // (`zz1`), while the BINDING keeps its namespace marker (`$zz1`).
            // Keying on the marked name is what made every php lookup miss.
            let key_text = self.variable_name_body(name).to_string();
            self.emit_global_read(GLOBAL_NAMESPACE_CACHE);
            self.class_get(
                crate::primitives::class_slots::ObjSource::Stack,
                &crate::primitives::class_slots::ClassSlot::internal("__keys"),
            );
            self.emit_const(Value::String(Arc::from(key_text.as_str())));
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        self.chunk().emit_end(line);

        // Values, every time. A dict entry set here is readable by name and,
        // because its key went into `__keys` above, visible to iteration.
        for name in &members {
            let key_text = self.variable_name_body(name).to_string();
            self.emit_global_read(GLOBAL_NAMESPACE_CACHE);
            self.emit_var_get(name);
            self.class_set(
                crate::primitives::class_slots::ObjSource::Stack,
                &crate::primitives::class_slots::ClassSlot::internal(&key_text),
                crate::primitives::class_slots::ValueSource::Stack,
            );
        }
        self.emit_global_read(GLOBAL_NAMESPACE_CACHE);
    }

    /// True when `expr` names this profile's global namespace AND that
    /// namespace has no materialized object to index.
    ///
    /// ECMA's global object is real and carries host-installed builtins, so
    /// `globalThis[k]` must keep resolving through it — routing it through the
    /// declared-member chain lost `globalThis.Math` and friends
    /// (`test_js_globalthis_builtin_constructors_accessible`). Lua `_G`,
    /// Python `globals()` and PHP `$GLOBALS` have no such object: `_G` on its
    /// own evaluates to nil today, which is why they need the chain.
    pub(super) fn expr_is_global_namespace(&self, expr: &Expression) -> bool {
        // ⛔ THE SPELLING IS GONE FROM SHARED CODE. This used to match the
        // source text against a `global_namespace` profile row (`_G` /
        // `$GLOBALS` / `globalThis` / `globals`) plus a
        // `global_namespace_is_call` flag for python's call form — a
        // per-language spelling table consulted by the shared compiler, and a
        // second home for a fact each walker already knew. All four walkers now
        // normalize their own word into `ExprKind::GlobalNamespace`; the rows
        // and both predicates are gone, and nothing here knows how any language
        // spells anything.
        //
        // ⚠ The `has_ecma_globals` early-return is the REMAINING asymmetry and
        // it is not a language difference: js has a real ECMA global object to
        // resolve against — and must keep resolving through it, or
        // `globalThis.Math` and the prelude's Map/Set wrappers go missing —
        // while lua/python/php have none, so they take the member chain below.
        // One place, not four rows.
        if self.profile.has_ecma_globals {
            return false;
        }
        matches!(&expr.kind, ExprKind::GlobalNamespace)
    }
}

/// Emit a read of module global `name` into a chunk.
///
/// **The one place a global access is encoded.** It used to be
/// `str_const(name)` + `emit_u16(GLOBAL_GET, idx)` open-coded at ~320 sites
/// across the shared primitives and the fifteen languages' emitter adapters.
/// Each site was free to intern its own constant for a name another site had
/// already interned, or to spell it differently — divergence whose symptom is
/// a global that reads `undefined` for no visible reason.
///
/// Language emitter adapters call this exactly as the primitives do: they are
/// emit layers, and this is what emitting a global means.
pub fn emit_read(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.intern_string_constant(name);
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

/// Emit a write of module global `name`. See [`emit_read`].
pub fn emit_write(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.intern_string_constant(name);
    chunk.emit_op_u16(Op::GLOBAL_SET, idx, line);
}

impl Compiler {
    /// Read module global `name` — the compiler-side entry point.
    ///
    /// Delegates to [`emit_read`] so there is ONE encoding of a global access,
    /// which is what this module claims to be.
    ///
    /// It used to route through a `global_name_const_idx` that returned either
    /// a shared-global SLOT number or a constant index, from the same function.
    /// `GLOBAL_GET` consumes a constant index, so the slot branch was only ever
    /// correct in chunk 0 — where the reserved names had been planted into an
    /// empty constant pool in slot order, making slot N land at constant N.
    /// In any other chunk it read an unrelated constant, or ran off the end of
    /// a short pool and panicked. Interning per chunk is correct everywhere and
    /// needs no coincidence.
    pub(crate) fn emit_global_read(&mut self, name: &str) {
        let line = self.line;
        emit_read(self.chunk(), name, line);
    }

    /// Write module global `name`. See [`Compiler::emit_global_read`].
    pub(crate) fn emit_global_write(&mut self, name: &str) {
        let line = self.line;
        emit_write(self.chunk(), name, line);
    }
}

/// Build the module's GLOBAL INDEX SPACE and rewrite every operand into it.
///
/// Emitters name a global; this assigns it an index. Modelled exactly on
/// `link.rs::normalize_import_table`, which does the same job for function
/// imports: unify across chunks, build a per-chunk remap, walk the bytecode,
/// rewrite the operand, park the table on chunk 0.
///
/// Before this pass a `GLOBAL_GET` operand was an index into the EMITTING
/// CHUNK's constant pool, holding the global's name, which the VM then looked
/// up in a `HashMap<String, Value>`. That is not what `global.get` means in
/// WASM and it cost a string hash on every access. After it, the operand is a
/// `globalidx` over `global_imports ++ defined` — WASM's own order, which
/// `collect_globals` in the wasm writer already assumes.
///
/// Runs AFTER `declare_free_globals`, which is what decides the import half.
pub fn normalize_global_table(chunks: &mut [Chunk]) {
    if chunks.is_empty() {
        return;
    }

    // The import halves, split exactly as the bytecode names them: a
    // string-constant global is referenced by its COMPOSITE key, a host global
    // by its BARE name.
    let mut string_constants: Vec<String> = Vec::new();
    let mut host_globals: Vec<String> = Vec::new();
    for chunk in chunks.iter() {
        for imp in &chunk.global_imports {
            if imp.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE {
                if !string_constants.contains(&imp.name) {
                    string_constants.push(imp.name.clone());
                }
            } else if !host_globals.contains(&imp.name) {
                host_globals.push(imp.name.clone());
            }
        }
    }

    // Everything else any chunk reads or writes, first-seen, is module-defined.
    let seeded = vybe_runtime::chunk::global_index_space(&string_constants, &host_globals, &[]);
    let mut defined: Vec<String> = Vec::new();
    for chunk in chunks.iter() {
        for (_, name) in global_operands(chunk) {
            if !seeded.contains(&name) && !defined.contains(&name) {
                defined.push(name);
            }
        }
    }

    let table =
        vybe_runtime::chunk::global_index_space(&string_constants, &host_globals, &defined);

    let mut remaps: Vec<Vec<(u16, u16)>> = Vec::with_capacity(chunks.len());
    for chunk in chunks.iter() {
        let mut remap: Vec<(u16, u16)> = Vec::new();
        for (const_idx, name) in global_operands(chunk) {
            let Some(gidx) = table.iter().position(|n| *n == name) else {
                continue;
            };
            if !remap.iter().any(|(c, _)| *c == const_idx) {
                remap.push((const_idx, gidx as u16));
            }
        }
        remaps.push(remap);
    }

    for (chunk_idx, chunk) in chunks.iter_mut().enumerate() {
        let remap = &remaps[chunk_idx];
        let code = &mut chunk.code;
        let mut ip = 0usize;
        while ip + 3 < code.len() {
            let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
            let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
            let Some(op) = Op::decode(group, sub) else {
                ip += 4;
                continue;
            };
            let operand_start = ip + 4;
            let operand_len = op.operand_format().size_in(code, operand_start);
            if (op == Op::GLOBAL_GET || op == Op::GLOBAL_SET) && operand_start + 1 < code.len() {
                let old = u16::from_be_bytes([code[operand_start], code[operand_start + 1]]);
                if let Some((_, gidx)) = remap.iter().find(|(c, _)| *c == old) {
                    let bytes = gidx.to_be_bytes();
                    code[operand_start] = bytes[0];
                    code[operand_start + 1] = bytes[1];
                }
            }
            ip = operand_start + operand_len;
        }
    }

    // Share the one table with EVERY chunk, not just chunk 0. The index space
    // is module-level, so any chunk holding a `GLOBAL_GET` operand can resolve
    // it — that is what lets `disassemble(chunk)` name globals without being
    // handed the table, and what keeps a single source of truth for the space.
    let shared = std::sync::Arc::new(table);
    for chunk in chunks.iter_mut() {
        chunk.globals = shared.clone();
    }
}

/// Every `(constant index, global name)` a chunk's `GLOBAL_GET`/`GLOBAL_SET`
/// operands name. Shared by `declare_free_globals` and
/// `normalize_global_table` so the two cannot disagree about what a global is.
fn global_operands(chunk: &Chunk) -> Vec<(u16, String)> {
    let code = &chunk.code;
    let mut out = Vec::new();
    let mut ip = 0usize;
    while ip + 3 < code.len() {
        let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
        let Some(op) = Op::decode(group, sub) else {
            ip += 4;
            continue;
        };
        let operand_start = ip + 4;
        if (op == Op::GLOBAL_GET || op == Op::GLOBAL_SET) && operand_start + 1 < code.len() {
            let idx = u16::from_be_bytes([code[operand_start], code[operand_start + 1]]);
            if let Some(vybe_runtime::Value::String(name)) = chunk.constants.get(idx as usize) {
                out.push((idx, name.to_string()));
            }
        }
        ip = operand_start + op.operand_format().size_in(code, operand_start);
    }
    out
}

/// Declare the module's FREE globals as imports.
///
/// Lives here because it is global-namespace logistics, which is what this
/// module is for. It reads the emitted bytecode rather than a record kept
/// during emission.
///
/// ⚠ The reason it gives used to be "only 16 of the 193 `GLOBAL_GET`/
/// `GLOBAL_SET` emit sites funnel through `global_name_const_idx`, so a
/// per-site record would be INCOMPLETE", with the note that funnelling every
/// site "is the real fix, and it is a separate sweep". **That sweep has since
/// been done** — every global emission in the compiler and all fifteen
/// languages now goes through `emit_read`/`emit_write` in this file, and
/// `global_name_const_idx` is gone. Measure before repeating the old figure:
/// it misled two sessions into pricing a four-function change as a
/// tree-wide one.
///
/// So a per-site record IS now possible and would let the record replace this
/// walk. It has not been done because the walk works and is not the bottleneck,
/// not because it cannot be.
///
/// A free global is a name the module reads but never writes and never
/// defines — `globalThis`, `undefined`, `__ctor_TypeError`, the runtime
/// helper anchors. In WASM a module may only touch globals it declared, so
/// these are imports, and saying so is what makes the module
/// self-describing instead of relying on whatever the host happens to have
/// left in a shared map.
///
/// Measured 2026-08-04: 7 free names for a Python class program, 6 for
/// PHP, 19 for JS. An import section, not a problem — the earlier estimate
/// of "hundreds" was wrong.
///
/// Computed from the emitted bytecode rather than threaded through the emit
/// sites, exactly as `normalize_import_table` walks the code for
/// `CALL_IMPORT`. String constants are skipped: they are already
/// declared imports of their own.
///
/// The test is "no chunk WRITES it", not `defined_globals`. Measured
/// 2026-08-04: `defined_globals` records what the SOURCE declares, and the
/// prelude declares `globalThis`, `undefined`, `Function`, `__ctor_Error`
/// — names whose values the host supplies and no chunk ever assigns.
/// Subtracting it emptied the set entirely. Whoever writes a global is its
/// definition; that is a fact about the bytecode.
pub fn declare_free_globals(chunks: &mut [Chunk]) {
    if chunks.is_empty() {
        return;
    }
    let mut read: Vec<String> = Vec::new();
    let mut written: HashSet<String> = HashSet::new();

    for chunk in chunks.iter() {
        let code = &chunk.code;
        let mut ip = 0usize;
        while ip + 3 < code.len() {
            let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
            let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
            let Some(op) = Op::decode(group, sub) else {
                ip += 4;
                continue;
            };
            let operand_start = ip + 4;
            if (op == Op::GLOBAL_GET || op == Op::GLOBAL_SET) && operand_start + 1 < code.len() {
                let idx =
                    u16::from_be_bytes([code[operand_start], code[operand_start + 1]]) as usize;
                if let Some(vybe_runtime::Value::String(name)) = chunk.constants.get(idx) {
                    let name = name.to_string();
                    if !name.starts_with(vybe_runtime::chunk::STRING_CONSTANTS_MODULE) {
                        if op == Op::GLOBAL_SET {
                            written.insert(name);
                        } else if !read.contains(&name) {
                            read.push(name);
                        }
                    }
                }
            }
            ip = operand_start + op.operand_format().size_in(code, operand_start);
        }
    }

    let mut declared = 0usize;
    for name in read {
        // A name this module writes is its own storage, not an import.
        if written.contains(&name) {
            continue;
        }
        chunks[0].add_global_import(vybe_runtime::chunk::HOST_GLOBALS_MODULE, &name);
        declared += 1;
    }
    if std::env::var("VYBE_DEBUG_IMPORTS").is_ok() {
        eprintln!(
            "[free-globals] {} declared, {} written, {} chunks",
            declared,
            written.len(),
            chunks.len()
        );
    }
}
