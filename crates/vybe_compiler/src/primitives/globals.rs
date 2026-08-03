//! The module's global namespace as ONE object.
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
