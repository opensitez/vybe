/// Compile-time scope tracking for local variable resolution.
/// (Identical in structure to the JS compiler scope.)

#[derive(Debug, Clone)]
pub struct Local {
    pub name: String,
    pub depth: u32,
    pub is_captured: bool,
    pub slot: u16,
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
    /// Names declared `global` in this function scope.
    pub globals: std::collections::HashSet<String>,
}

impl Scope {
    pub fn new() -> Self {
        Scope {
            locals: Vec::new(),
            upvalues: Vec::new(),
            depth: 0,
            next_slot: 0,
            globals: std::collections::HashSet::new(),
        }
    }

    pub fn new_function() -> Self {
        let mut scope = Self::new();
        // Slot 0 is reserved for the function itself (callee).
        scope.locals.push(Local {
            name: String::new(),
            depth: 0,
            is_captured: false,
            slot: 0,
        });
        scope.next_slot = 1;
        scope
    }

    pub fn define_local(&mut self, name: &str) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local {
            name: name.to_string(),
            depth: self.depth,
            is_captured: false,
            slot,
        });
        self.next_slot += 1;
        slot
    }

    pub fn resolve_local(&self, name: &str) -> Option<u16> {
        for local in self.locals.iter().rev() {
            if local.name == name {
                return Some(local.slot);
            }
        }
        None
    }

    pub fn mark_captured(&mut self, slot: u16) {
        for local in &mut self.locals {
            if local.slot == slot {
                local.is_captured = true;
                return;
            }
        }
    }

    pub fn begin_block(&mut self) {
        self.depth += 1;
    }

    pub fn end_block(&mut self) {
        self.depth -= 1;
        self.locals.retain(|l| l.depth <= self.depth);
    }
}
