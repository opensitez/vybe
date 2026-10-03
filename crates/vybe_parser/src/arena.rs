//! Append-only semantic arenas and transactional value stacks. Speculation
//! journals small handles rather than cloning owned AST subtrees.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Id(usize);
enum Undo {
    RemoveNode,
    RemoveValue,
    RestoreValue(Id),
}
pub struct Arena<T> {
    nodes: Vec<T>,
    values: Vec<Id>,
    undo: Vec<Undo>,
}
impl<T> Default for Arena<T> {
    fn default() -> Self {
        Self {
            nodes: Vec::new(),
            values: Vec::new(),
            undo: Vec::new(),
        }
    }
}
impl<T> Arena<T> {
    pub fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    /// Marks/IDs belong to this arena. Speculative IDs must not escape hooks;
    /// rolling back invalidates allocations after the mark.
    pub fn rollback(&mut self, mark: usize) {
        assert!(mark <= self.undo.len(), "invalid semantic arena checkpoint");
        while self.undo.len() > mark {
            match self.undo.pop().unwrap() {
                Undo::RemoveNode => {
                    self.nodes.pop().unwrap();
                }
                Undo::RemoveValue => {
                    self.values.pop().unwrap();
                }
                Undo::RestoreValue(value) => self.values.push(value),
            }
        }
    }
    pub fn alloc(&mut self, node: T) -> Id {
        let id = Id(self.nodes.len());
        self.nodes.push(node);
        self.undo.push(Undo::RemoveNode);
        id
    }
    pub fn get(&self, id: Id) -> &T {
        &self.nodes[id.0]
    }
    pub fn push(&mut self, id: Id) {
        assert!(id.0 < self.nodes.len(), "invalid semantic node ID");
        self.values.push(id);
        self.undo.push(Undo::RemoveValue);
    }
    pub fn pop(&mut self) -> Option<Id> {
        let id = self.values.pop()?;
        self.undo.push(Undo::RestoreValue(id));
        Some(id)
    }
    pub fn values(&self) -> &[Id] {
        &self.values
    }
    pub fn node_count(&self) -> usize {
        self.nodes.len()
    }
    /// Release undo storage after a parse has committed and no old checkpoints
    /// will be used. Nodes/value IDs remain valid; old marks are invalidated.
    pub fn commit(&mut self) {
        self.undo.clear();
    }
}
