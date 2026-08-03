/// Scope tracking for local variables and upvalue captures.

#[derive(Debug, Clone)]
pub struct Local {
    pub name: String,
    pub depth: u32,
    pub slot: u16,
    pub is_captured: bool,
    pub type_hint: Option<String>,
    /// `true` for bindings introduced by `const` (ECMA-262 immutable
    /// bindings). Reassigning one is a runtime `TypeError`. Only the
    /// `emit_var_set` assignment path consults this — declaration init
    /// and direct loop-variable rebinds use `LOCAL_SET` directly, so
    /// they are unaffected.
    pub is_const: bool }

#[derive(Debug, Clone)]
pub struct UpvalueDesc {
    /// For `is_local`: the parent frame's LOCAL SLOT (u16 to match
    /// `Local.slot` — chunks routinely exceed 255 locals). Otherwise:
    /// the parent's upvalue-list position.
    pub index: u16,
    pub is_local: bool }

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
    Closed }

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
    /// True for vb, pascal, cobol and fortran; false everywhere else.
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
    /// this covers locals only. Callable/type folding is
    /// `LanguageProfile::fold_callable_names`.
    pub fold_case: bool,
    pub depth: u32,
    pub next_slot: u16,
    /// Debug accumulator: every `(slot, name)` ever defined in this function,
    /// NOT popped by `end_scope`. Copied into `Chunk.local_names` at finalize
    /// so the debugger can resolve variable names ↔ slots. Inspection only.
    pub defined_names: Vec<(u16, String)> }

impl Scope {
    /// `fold_case` is required rather than defaulted: a scope that silently
    /// stopped folding would mis-resolve every vb/pascal/cobol/fortran local,
    /// and the compiler cannot catch that. Passing it makes a missed
    /// construction site a build error instead.
    pub fn new(fold_case: bool) -> Self {
        Self {
            locals: Vec::new(),
            upvalues: Vec::new(),
            resolution: ScopeResolution::Chain,
            open_names: std::collections::HashSet::new(),
            fold_case,
            depth: 0,
            next_slot: 0,
            defined_names: Vec::new() }
    }

    pub fn new_function(fold_case: bool) -> Self {
        // WASM convention: slot 0 is the first argument (not a reserved callee).
        // User-visible locals (params and additional locals) start at slot 0.
        Self::new(fold_case)
    }

    /// A function scope that inherits the enclosing scope's resolution.
    ///
    /// A closure resolves names the way the body containing it does — a PHP
    /// closure sees no more of the module than the function it sits in. Its own
    /// `open_names` start empty: `global $x;` binds one function, not the
    /// lambdas nested inside it.
    pub fn new_function_like(enclosing: ScopeResolution, fold_case: bool) -> Self {
        Scope {
            resolution: enclosing,
            ..Self::new_function(fold_case)
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

    pub fn define(&mut self, name: &str) -> u16 {
        self.define_typed(name, None)
    }

    pub fn define_typed(&mut self, name: &str, type_hint: Option<String>) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local {
            name: name.to_string(),
            depth: self.depth,
            slot,
            is_captured: false,
            type_hint,
            is_const: false });
        self.defined_names.push((slot, name.to_string()));
        self.next_slot += 1;
        slot
    }

    /// Define a local at the function-scope depth (depth = 0) regardless
    /// of the current block depth. Used for `var` declarations in JS,
    /// which are function-scoped (ECMA-262 §10.2.11) — the binding must
    /// outlive the block it was declared in. Without this, `end_scope`
    /// would pop the binding when the enclosing block exits.
    pub fn define_at_function_scope(&mut self, name: &str, type_hint: Option<String>) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local {
            name: name.to_string(),
            depth: 0,
            slot,
            is_captured: false,
            type_hint,
            is_const: false });
        self.defined_names.push((slot, name.to_string()));
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
        for l in self.locals.iter().rev() {
            if l.name == name {
                return Some(l.slot);
            }
        }
        if self.fold_case {
            for l in self.locals.iter().rev() {
                if l.name.eq_ignore_ascii_case(name) {
                    return Some(l.slot);
                }
            }
        }
        None
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
        if self.fold_case {
            for l in self.locals.iter().rev() {
                if l.name.eq_ignore_ascii_case(name) {
                    return l.type_hint.as_deref();
                }
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
