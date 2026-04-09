/// Scope tracking for local variables and upvalue captures.

#[derive(Debug, Clone)]
pub struct Local {
    pub name: String,
    pub depth: u32,
    pub slot: u16,
    pub is_captured: bool,
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
        Self { locals: Vec::new(), upvalues: Vec::new(), depth: 0, next_slot: 0 }
    }

    pub fn new_function() -> Self {
        let mut s = Self::new();
        s.locals.push(Local { name: String::new(), depth: 0, slot: 0, is_captured: false });
        s.next_slot = 1;
        s
    }

    pub fn define(&mut self, name: &str) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local { name: name.to_string(), depth: self.depth, slot, is_captured: false });
        self.next_slot += 1;
        slot
    }

    pub fn resolve(&self, name: &str) -> Option<u16> {
        for l in self.locals.iter().rev() {
            if l.name == name { return Some(l.slot); }
        }
        None
    }

    pub fn resolve_ci(&self, name: &str) -> Option<u16> {
        for l in self.locals.iter().rev() {
            if l.name.eq_ignore_ascii_case(name) { return Some(l.slot); }
        }
        None
    }

    pub fn mark_captured(&mut self, slot: u16) {
        for l in &mut self.locals { if l.slot == slot { l.is_captured = true; return; } }
    }

    pub fn begin_scope(&mut self) { self.depth += 1; }

    pub fn end_scope(&mut self) {
        while let Some(l) = self.locals.last() {
            if l.depth < self.depth { break; }
            self.locals.pop();
        }
        self.depth -= 1;
    }

    pub fn add_upvalue(&mut self, index: u8, is_local: bool) -> u8 {
        for (i, uv) in self.upvalues.iter().enumerate() {
            if uv.index == index && uv.is_local == is_local { return i as u8; }
        }
        let idx = self.upvalues.len() as u8;
        self.upvalues.push(UpvalueDesc { index, is_local });
        idx
    }
}
