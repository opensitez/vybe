//! The module's global namespace as ONE object — and the logistics that go
//! with it: which names the module OWNS, and which it merely reads.
//!
//! Four languages spell the same thing:
//!
//! | language | spelling | form |
//! |---|---|---|
//! | Lua | `_G` | identifier |
//! | JS | `globalThis` | identifier |
//! | PHP | `$GLOBALS` | identifier |
//! | Python | `globals()` | zero-argument call |
//!
//! All four name the module's OBJECT environment record — the same storage a
//! bare `x` reads. `vm.globals` is already a live `HashMap<String, Value>`
//! (`vm.rs:552`), and `GLOBAL_GET` is nothing but "resolve a constant index to
//! a string, look it up in that map". So a namespace object is a *view* over
//! storage that already exists; nothing new has to be stored.
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
    pub lexical_bindings_are_members: bool }

impl Options {
    pub fn from_profile(profile: &LanguageProfile) -> Options {
        Options {
            lexical_bindings_are_members: !profile.has_ecma_globals }
    }
}

/// True when `name` is this profile's spelling of the global namespace, in
/// IDENTIFIER form (`_G`, `globalThis`, `$GLOBALS`).
pub fn names_global_namespace(profile: &LanguageProfile, name: &str) -> bool {
    !profile.global_namespace.is_empty()
        && !profile.global_namespace_is_call
        && profile.global_namespace == name
}

/// True when `name` is this profile's spelling in CALL form — Python's
/// `globals()`. The caller has already checked the argument list is empty.
pub fn names_global_namespace_call(profile: &LanguageProfile, name: &str) -> bool {
    profile.global_namespace_is_call && profile.global_namespace == name
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
                    // not user bindings and must not be exposed.
                    .filter(|(_, name)| !name.starts_with("__"))
                    .map(|(_, name)| name.clone()),
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
        if self.profile.has_ecma_globals {
            return false;
        }
        match &expr.kind {
            ExprKind::Ident(n) => names_global_namespace(&self.profile, n),
            ExprKind::Call { callee, args, .. } if args.is_empty() => {
                matches!(&callee.kind, ExprKind::Ident(n)
                    if names_global_namespace_call(&self.profile, n))
            }
            _ => false }
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
    /// Routes through `global_name_const_idx` so the shared-global slot map is
    /// honoured. That map existed while only **16** of the ~320 emit sites
    /// consulted it; every other site interned its own constant and could
    /// refer to a different slot for the same name.
    pub(crate) fn emit_global_read(&mut self, name: &str) {
        let idx = self.global_name_const_idx(name);
        self.emit_u16(Op::GLOBAL_GET, idx);
    }

    /// Write module global `name`. See [`Compiler::emit_global_read`].
    pub(crate) fn emit_global_write(&mut self, name: &str) {
        let idx = self.global_name_const_idx(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
    }
}

/// Declare the module's FREE globals as imports.
///
/// Lives here because it is global-namespace logistics, which is what this
/// module is for. It reads the emitted bytecode rather than a record kept
/// during emission for one reason, stated plainly: only 16 of the 193
/// `GLOBAL_GET`/`GLOBAL_SET` emit sites funnel through
/// `global_name_const_idx`, so a per-site record would be INCOMPLETE and an
/// incomplete import list is worse than none. Funnelling every global read and
/// write through this module would let the record replace the walk — that is
/// the real fix, and it is a separate sweep.
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
/// Computed from the emitted bytecode rather than threaded through the 193
/// `GLOBAL_GET` emit sites, exactly as `normalize_import_table` walks the
/// code for `CALL_IMPORT`. String constants are skipped: they are already
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
            if (op == Op::GLOBAL_GET || op == Op::GLOBAL_SET) && operand_start + 1 < code.len()
            {
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
