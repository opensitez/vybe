/// Scope tracking for local variables and upvalue captures.

#[derive(Debug, Clone)]
pub struct Local {
    pub name: String,
    pub depth: u32,
    pub slot: u16,
    pub is_captured: bool,
    pub type_hint: Option<vybe_ast::TypeHint>,
    /// `true` for bindings introduced by `const` (ECMA-262 immutable
    /// bindings). Reassigning one is a runtime `TypeError`. Only the
    /// `emit_var_set` assignment path consults this — declaration init
    /// and direct loop-variable rebinds use `LOCAL_SET` directly, so
    /// they are unaffected.
    pub is_const: bool,
    /// This binding holds a REFERENCE (`{__ref_kind:"cell"}` or `"carray"`), not
    /// a value: reads auto-deref and writes store through to whatever it points
    /// at. Set for a `PassBy::Alias` parameter, for `$r = &$x`, and for a local
    /// promoted by having its address taken.
    ///
    /// It lives on the BINDING for the same reason `is_const` does. It used to
    /// be a `HashMap<chunk_idx, HashSet<name>>` beside the compiler, which asked
    /// a name-keyed side table a question only the resolver can answer: a name
    /// marked in ANY chunk leaked into every other chunk through a module-wide
    /// fallback, and the guard added to stop that ("this chunk has a local of
    /// that name → not a cell") could not tell a local that SHADOWS the name
    /// from one that IS it — so php's `global $g` read a promoted global raw.
    /// Resolution already knows the difference. Ask it.
    pub holds_reference: bool,
    /// This binding is a FUNCTION DECLARED HERE — a nested `function`/
    /// `procedure` — and the value is its parameter count.
    ///
    /// It lives on the binding for the same reason `is_const` and
    /// `holds_reference` do. A nested declaration deliberately does NOT enter
    /// `defined_functions` (`classes.rs` — a flat set made sibling frames race
    /// for one name), so the only record that a local IS a callable was gone by
    /// the time an expression asked. `function_min_arity` could not answer
    /// either: it is a flat name-keyed map, so a local variable sharing a name
    /// with some unrelated function elsewhere would answer for it.
    /// Resolution knows which binding is in scope. Ask it.
    pub declared_arity: Option<u8>,
}

#[derive(Debug, Clone)]
pub struct UpvalueDesc {
    /// For `is_local`: the parent frame's LOCAL SLOT (u16 to match
    /// `Local.slot` — chunks routinely exceed 255 locals). Otherwise:
    /// the parent's upvalue-list position.
    pub index: u16,
    pub is_local: bool,
}

/// What a name lookup that misses this scope's own locals does.
///
/// This is a property of the SCOPE, not of a language: `Scope::new_function`
/// already existed as a distinct constructor but carried no behaviour, so the
/// one language whose function scopes differ (PHP) had the rule spread across
/// five `profile.name == "php"` checks and a `php_function_globals` field on
/// the shared compiler. It lives here now, and languages declare it through
/// `ScopeDeclKind::Closed` like any other scope statement.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ScopeResolution {
    /// A miss falls through to the enclosing/module scope. Every language
    /// except PHP, and PHP's own module scope — the behaviour before this
    /// existed, hence the default.
    #[default]
    Chain,
    /// A miss resolves to nothing: reads are null, assignments create a local
    /// here. PHP function bodies.
    Closed,
}

#[derive(Debug)]
pub struct Scope {
    pub locals: Vec<Local>,
    pub upvalues: Vec<UpvalueDesc>,
    /// See [`ScopeResolution`].
    pub resolution: ScopeResolution,
    /// Names re-opened to an outer scope regardless of `resolution` — PHP
    /// `global $x;` and its superglobals, Python `global`/`nonlocal`. Empty
    /// under [`ScopeResolution::Chain`], where everything is open anyway.
    pub open_names: std::collections::HashSet<String>,
    /// Whether two names differing only in ASCII case are the SAME name here.
    /// `None` folds nothing; `Some(alphabet)` folds over that alphabet.
    /// Stated by the module's `Directives::variable_case` — vb, pascal, cobol,
    /// fortran and powershell say `Folded`; everything else says nothing and
    /// gets `Exact`.
    ///
    /// This is a resolution POLICY and belongs beside `resolution`, for the
    /// same reason: it used to be a bare `resolve_ci` scan with the
    /// `case_sensitive` flag living on the compiler, so all 33 call sites had
    /// to remember to write `!self.case_sensitive &&` themselves. 23 of them
    /// did not, and two of those silently broke go — a local `ab` matched the
    /// class `AB`, so `ab.Get()` stopped resolving as a class reference and
    /// compiled to a call that passed no receiver.
    ///
    /// Folding is per NAME KIND, not per language: PHP variables are
    /// case-sensitive while its function and class names are not, which is why
    /// this covers locals only and why `Directives` carries `variable_case`
    /// and `callable_case` as two fields.
    ///
    /// ⛔ It carries the ALPHABET, not a bool, because the two folds already in
    /// this pipeline disagree: here it is ASCII, in
    /// `vybe_runtime::namespaces` it is Unicode `to_lowercase()`. A bool cannot
    /// state which one a language meant.
    pub fold: Option<vybe_ast::CaseAlphabet>,
    pub depth: u32,
    pub next_slot: u16,
    /// Debug accumulator: every `(slot, name)` ever defined in this function,
    /// NOT popped by `end_scope`. Copied into `Chunk.local_names` at finalize
    /// so the debugger can resolve variable names ↔ slots. Inspection only.
    pub defined_names: Vec<vybe_runtime::chunk::LocalName>,
    /// Origin stamped on the next binding defined; reset to `Compiler` after
    /// each one so a forgotten `set_pending_origin` cannot leak a Source mark
    /// onto an unrelated temporary.
    pending_origin: vybe_runtime::chunk::LocalOrigin,
}

impl Scope {
    /// `fold` is required rather than defaulted: a scope that silently stopped
    /// folding would mis-resolve every vb/pascal/cobol/fortran/powershell
    /// local, and the compiler cannot catch that. Passing it makes a missed
    /// construction site a build error instead. Keep it that way.
    pub fn new(fold: Option<vybe_ast::CaseAlphabet>) -> Self {
        Self {
            locals: Vec::new(),
            upvalues: Vec::new(),
            resolution: ScopeResolution::Chain,
            open_names: std::collections::HashSet::new(),
            fold,
            depth: 0,
            next_slot: 0,
            defined_names: Vec::new(),
            pending_origin: vybe_runtime::chunk::LocalOrigin::Compiler,
        }
    }

    pub fn new_function(fold: Option<vybe_ast::CaseAlphabet>) -> Self {
        // WASM convention: slot 0 is the first argument (not a reserved callee).
        // User-visible locals (params and additional locals) start at slot 0.
        Self::new(fold)
    }

    /// A function scope that inherits the enclosing scope's resolution.
    ///
    /// A closure resolves names the way the body containing it does — a PHP
    /// closure sees no more of the module than the function it sits in. Its own
    /// `open_names` start empty: `global $x;` binds one function, not the
    /// lambdas nested inside it.
    pub fn new_function_like(
        enclosing: ScopeResolution,
        fold: Option<vybe_ast::CaseAlphabet>,
    ) -> Self {
        Scope {
            resolution: enclosing,
            ..Self::new_function(fold)
        }
    }

    /// True when `name` may resolve outside this scope's own locals.
    ///
    /// Under [`ScopeResolution::Chain`] everything may, which is why every
    /// language but PHP is unaffected by this whole mechanism.
    pub fn is_open(&self, name: &str) -> bool {
        self.resolution == ScopeResolution::Chain || self.open_names.contains(name)
    }

    /// Apply a `ScopeDeclKind::Closed` declaration: close the scope and seed
    /// the names that stay open regardless (PHP's superglobals).
    pub fn close(&mut self, always_open: &[String]) {
        self.resolution = ScopeResolution::Closed;
        self.open_names.extend(always_open.iter().cloned());
    }

    /// Apply a `global` / `nonlocal` declaration.
    pub fn open(&mut self, names: &[String]) {
        self.open_names.extend(names.iter().cloned());
    }

    /// Was this name EXPLICITLY re-opened here (`global $g`, `nonlocal x`)?
    ///
    /// Distinct from [`Scope::is_open`], which is also true for every name under
    /// [`ScopeResolution::Chain`]. The difference matters for per-binding
    /// properties: a declared-open name IS the outer binding, so any local
    /// record for it aliases rather than shadows, and its properties must be
    /// read from where the binding actually lives.
    pub fn declared_open(&self, name: &str) -> bool {
        self.open_names.contains(name)
    }

    pub fn define(&mut self, name: &str) -> u16 {
        self.define_typed(name, None)
    }

    /// Record who declared the NEXT binding this scope defines.
    ///
    /// The origin is a property of the call site, not of the name, and the
    /// call sites outnumber the definitions 909 to 2 — so it is set here and
    /// consumed by `define_typed` rather than threaded through every overload.
    pub fn set_pending_origin(&mut self, origin: vybe_runtime::chunk::LocalOrigin) {
        self.pending_origin = origin;
    }

    pub fn define_typed(&mut self, name: &str, type_hint: Option<vybe_ast::TypeHint>) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local {
            name: name.to_string(),
            depth: self.depth,
            slot,
            is_captured: false,
            type_hint,
            is_const: false,
            holds_reference: false,
            declared_arity: None,
        });
        self.defined_names.push(vybe_runtime::chunk::LocalName::new(
            slot,
            name,
            std::mem::replace(
                &mut self.pending_origin,
                vybe_runtime::chunk::LocalOrigin::Compiler,
            ),
        ));
        self.next_slot += 1;
        slot
    }

    /// Define a local at the function-scope depth (depth = 0) regardless
    /// of the current block depth. Used for `var` declarations in JS,
    /// which are function-scoped (ECMA-262 §10.2.11) — the binding must
    /// outlive the block it was declared in. Without this, `end_scope`
    /// would pop the binding when the enclosing block exits.
    pub fn define_at_function_scope(
        &mut self,
        name: &str,
        type_hint: Option<vybe_ast::TypeHint>,
    ) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local {
            name: name.to_string(),
            depth: 0,
            slot,
            is_captured: false,
            type_hint,
            is_const: false,
            holds_reference: false,
            declared_arity: None,
        });
        self.defined_names.push(vybe_runtime::chunk::LocalName::new(
            slot,
            name,
            std::mem::replace(
                &mut self.pending_origin,
                vybe_runtime::chunk::LocalOrigin::Compiler,
            ),
        ));
        self.next_slot += 1;
        slot
    }

    /// Mark the most-recently-defined local with the given slot as a
    /// `const` binding (see `Local::is_const`).
    pub fn mark_const(&mut self, slot: u16) {
        for l in self.locals.iter_mut().rev() {
            if l.slot == slot {
                l.is_const = true;
                return;
            }
        }
    }

    /// Record that the local in `slot` is a nested function of `arity`
    /// parameters (see `Local::declared_arity`).
    pub fn mark_declared_arity(&mut self, slot: u16, arity: u8) {
        for l in self.locals.iter_mut().rev() {
            if l.slot == slot {
                l.declared_arity = Some(arity);
                return;
            }
        }
    }

    /// The parameter count of the nested function bound to `name` here, or
    /// `None` when `name` is not in this scope or is an ordinary value.
    pub fn resolve_declared_arity(&self, name: &str) -> Option<u8> {
        self.resolve_local(name).and_then(|l| l.declared_arity)
    }

    /// Returns `true` if a binding for `name` is in scope and is `const`.
    pub fn resolve_is_const(&self, name: &str) -> bool {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return l.is_const;
            }
        }
        false
    }

    /// Exact match first, THEN a folded pass — never one folded scan.
    ///
    /// Every call site this replaced was exact-then-folded (`resolve()
    /// .or_else(resolve_ci)`, or an `if !case_sensitive` block after the exact
    /// lookup returned), so two passes reproduce all of them. A single folded
    /// scan would not: the compiler defines its own locals beside the user's,
    /// so a scope can hold both `Result` (pascal's `result_slot_name`) and a
    /// user's `result`, and one pass would return whichever sits later.
    pub fn resolve(&self, name: &str) -> Option<u16> {
        self.resolve_local(name).map(|l| l.slot)
    }

    /// The binding `resolve` would pick — innermost first, then case-folded if
    /// this scope folds. Every per-binding property is read through here, so it
    /// can never disagree with the slot `resolve` returns.
    fn resolve_local(&self, name: &str) -> Option<&Local> {
        if let Some(l) = self.locals.iter().rev().find(|l| l.name == name) {
            return Some(l);
        }
        let fold = self.fold?;
        self.locals
            .iter()
            .rev()
            .find(|l| names_equal(fold, &l.name, name))
    }

    fn resolve_local_mut(&mut self, name: &str) -> Option<&mut Local> {
        if self.locals.iter().rev().any(|l| l.name == name) {
            return self.locals.iter_mut().rev().find(|l| l.name == name);
        }
        let fold = self.fold?;
        self.locals
            .iter_mut()
            .rev()
            .find(|l| names_equal(fold, &l.name, name))
    }

    /// Does this name denote a binding HERE that holds a reference?
    ///
    /// `None` means the name is not a local of this scope at all, so the caller
    /// must look further out — that is the distinction a name-keyed side table
    /// could not make, and the whole reason this lives on the binding.
    pub fn holds_reference(&self, name: &str) -> Option<bool> {
        self.resolve_local(name).map(|l| l.holds_reference)
    }

    /// Mark this scope's binding for `name` as holding a reference. Returns
    /// `false` when the name is not a local here, so the caller can record it
    /// as a promoted GLOBAL instead.
    pub fn set_holds_reference(&mut self, name: &str) -> bool {
        match self.resolve_local_mut(name) {
            Some(l) => {
                l.holds_reference = true;
                true
            }
            None => false,
        }
    }

    /// Case-EXACT, whatever the folding policy.
    ///
    /// For the one decision that legitimately distinguishes an exact match from
    /// a folded one: a name that only matches when folded loses to a
    /// type-qualified interpretation (`expressions.rs`). Everything else wants
    /// [`Scope::resolve`] — reach for this only when "did it match exactly?" is
    /// itself the question.
    pub fn resolve_exact(&self, name: &str) -> Option<u16> {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return Some(l.slot);
            }
        }
        None
    }

    /// Exact-then-folded, for the same reason as [`Scope::resolve`].
    pub fn resolve_type(&self, name: &str) -> Option<&str> {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return l.type_hint.as_deref();
            }
        }
        let fold = self.fold?;
        for l in self.locals.iter().rev() {
            if names_equal(fold, &l.name, name) {
                return l.type_hint.as_deref();
            }
        }
        None
    }

    /// The full declared type, not just its spelling — the caller needs
    /// [`vybe_ast::TypeBinding`], which `resolve_type` drops on the way to
    /// `&str`. Same exact-then-folded order as [`Scope::resolve_type`].
    pub fn resolve_declared(&self, name: &str) -> Option<&vybe_ast::TypeHint> {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return l.type_hint.as_ref();
            }
        }
        let fold = self.fold?;
        for l in self.locals.iter().rev() {
            if names_equal(fold, &l.name, name) {
                return l.type_hint.as_ref();
            }
        }
        None
    }

    pub fn mark_captured(&mut self, slot: u16) {
        for l in &mut self.locals {
            if l.slot == slot {
                l.is_captured = true;
                return;
            }
        }
    }

    pub fn begin_scope(&mut self) {
        self.depth += 1;
    }

    pub fn end_scope(&mut self) {
        while let Some(l) = self.locals.last() {
            if l.depth < self.depth {
                break;
            }
            self.locals.pop();
        }
        self.depth -= 1;
    }

    pub fn add_upvalue(&mut self, index: u16, is_local: bool) -> u8 {
        for (i, uv) in self.upvalues.iter().enumerate() {
            if uv.index == index && uv.is_local == is_local {
                return i as u8;
            }
        }
        let idx = self.upvalues.len() as u8;
        self.upvalues.push(UpvalueDesc { index, is_local });
        idx
    }
}

/// Compare a DECLARED identifier against a REFERENCE under a stated alphabet.
///
/// ⛔ ONE function, so the alphabets cannot drift apart the way this pipeline's
/// two already have: scopes fold ASCII (`eq_ignore_ascii_case`) while
/// `vybe_runtime::namespaces` folds Unicode (`to_lowercase`), in the same
/// program, with nothing comparing the two to notice.
///
/// Argument order is `declared, reference` and not the reverse because that is
/// the direction normalisation runs: the answer is the DECLARED spelling, which
/// is what the reference is rewritten to.
fn names_equal(fold: vybe_ast::CaseAlphabet, declared: &str, reference: &str) -> bool {
    match fold {
        vybe_ast::CaseAlphabet::Ascii => declared.eq_ignore_ascii_case(reference),
        vybe_ast::CaseAlphabet::Unicode => declared.to_lowercase() == reference.to_lowercase(),
    }
}

/// The folding policy a profile implies, for the ONE moment a module is not yet
/// in hand: `Compiler::with_profile` builds a root scope at construction, before
/// `compile_with_imports` has seen `Module.directives`.
///
/// ⛔ Every other site reads the DIRECTIVE. This exists so the root scope is not
/// silently non-folding for the instant before the module arrives; the module
/// then overwrites it. If the two ever disagree the directive wins, because it
/// is the one that knows which unit of a multi-language bundle is compiling.
pub fn profile_variable_fold(
    profile: &vybe_runtime::profile::LanguageProfile,
) -> Option<vybe_ast::CaseAlphabet> {
    if profile.case_sensitive {
        None
    } else {
        Some(vybe_ast::CaseAlphabet::Ascii)
    }
}

/// The namespace-tree lookup policy this module declares.
///
/// ⛔ ONE accessor, so no call site decides for itself whether the tree may
/// fold. That is the same rule `Directives::variable_fold()` follows for
/// scopes, and the reason is the one this codebase already paid for: when
/// `case_sensitive` was a flag every site had to consult, 23 of 33 forgot.
///
/// It reads `callable_fold()` rather than `variable_fold()` because a tree
/// lookup resolves a TYPE or a MEMBER, never a local — and PHP folds its
/// callables while its variables are exact.
pub fn tree_fold(d: &vybe_ast::Directives) -> vybe_runtime::namespaces::Fold {
    d.callable_fold()
}
