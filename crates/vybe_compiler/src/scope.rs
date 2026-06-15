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
    pub is_const: bool,
}

#[derive(Debug, Clone)]
pub struct UpvalueDesc {
    pub index: u8,
    pub is_local: bool,
}

#[derive(Debug)]
pub struct Scope {
    pub locals: Vec<Local>,
    pub upvalues: Vec<UpvalueDesc>,
    pub depth: u32,
    pub next_slot: u16,
}

impl Scope {
    pub fn new() -> Self {
        Self {
            locals: Vec::new(),
            upvalues: Vec::new(),
            depth: 0,
            next_slot: 0,
        }
    }

    pub fn new_function() -> Self {
        // WASM convention: slot 0 is the first argument (not a reserved callee).
        // User-visible locals (params and additional locals) start at slot 0.
        Self::new()
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
            is_const: false,
        });
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
            is_const: false,
        });
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

    pub fn resolve(&self, name: &str) -> Option<u16> {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return Some(l.slot);
            }
        }
        None
    }

    pub fn resolve_ci(&self, name: &str) -> Option<u16> {
        for l in self.locals.iter().rev() {
            if l.name.eq_ignore_ascii_case(name) {
                return Some(l.slot);
            }
        }
        None
    }

    pub fn resolve_type(&self, name: &str) -> Option<&str> {
        for l in self.locals.iter().rev() {
            if l.name == name {
                return l.type_hint.as_deref();
            }
        }
        None
    }

    pub fn resolve_type_ci(&self, name: &str) -> Option<&str> {
        for l in self.locals.iter().rev() {
            if l.name.eq_ignore_ascii_case(name) {
                return l.type_hint.as_deref();
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

    pub fn add_upvalue(&mut self, index: u8, is_local: bool) -> u8 {
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
