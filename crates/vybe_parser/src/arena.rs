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

/// Transactional ownership for tree-shaped, owned AST types. Combining nodes
/// moves children into their parent. An inverse function extracts them again
/// on rollback, so speculative construction needs neither Clone nor a second
/// materialization walk. Children must have exclusive ownership; consumed IDs
/// cannot be accessed until their composition has been rolled back.
pub struct OwnedArena<T> {
    nodes: Vec<Option<T>>,
    values: Vec<Id>,
    undo: Vec<OwnedUndo<T>>,
}
enum OwnedUndo<T> {
    RemoveNode,
    RemoveValue,
    RestoreValue(Id),
    SplitUnary {
        child: Id,
        split: fn(T) -> T,
    },
    SplitBinary {
        left: Id,
        right: Id,
        split: fn(T) -> (T, T),
    },
}
impl<T> Default for OwnedArena<T> {
    fn default() -> Self {
        Self {
            nodes: Vec::new(),
            values: Vec::new(),
            undo: Vec::new(),
        }
    }
}
impl<T> OwnedArena<T> {
    pub fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    pub fn alloc(&mut self, node: T) -> Id {
        let id = Id(self.nodes.len());
        self.nodes.push(Some(node));
        self.undo.push(OwnedUndo::RemoveNode);
        id
    }
    pub fn get(&self, id: Id) -> &T {
        self.nodes[id.0].as_ref().expect("consumed semantic node")
    }
    pub fn push(&mut self, id: Id) {
        assert!(self.nodes[id.0].is_some(), "consumed semantic node");
        self.values.push(id);
        self.undo.push(OwnedUndo::RemoveValue);
    }
    pub fn pop(&mut self) -> Option<Id> {
        let id = self.values.pop()?;
        self.undo.push(OwnedUndo::RestoreValue(id));
        Some(id)
    }
    pub fn values(&self) -> &[Id] {
        &self.values
    }
    /// `split` must be the inverse of `make`, preserving the exact child.
    /// Constructors/inverses must not panic or mutate external state.
    pub fn unary(&mut self, child: Id, make: impl FnOnce(T) -> T, split: fn(T) -> T) -> Id {
        let value = self.nodes[child.0].take().expect("consumed semantic node");
        let id = Id(self.nodes.len());
        self.nodes.push(Some(make(value)));
        self.undo.push(OwnedUndo::SplitUnary { child, split });
        id
    }
    /// `split` must recover the original left/right children in that order.
    pub fn binary(
        &mut self,
        left: Id,
        right: Id,
        make: impl FnOnce(T, T) -> T,
        split: fn(T) -> (T, T),
    ) -> Id {
        assert_ne!(left, right, "owned AST children cannot alias");
        assert!(
            self.nodes[left.0].is_some() && self.nodes[right.0].is_some(),
            "consumed semantic node"
        );
        let left_value = self.nodes[left.0].take().unwrap();
        let right_value = self.nodes[right.0].take().unwrap();
        let id = Id(self.nodes.len());
        self.nodes.push(Some(make(left_value, right_value)));
        self.undo
            .push(OwnedUndo::SplitBinary { left, right, split });
        id
    }
    pub fn rollback(&mut self, mark: usize) {
        assert!(mark <= self.undo.len(), "invalid semantic arena checkpoint");
        while self.undo.len() > mark {
            match self.undo.pop().unwrap() {
                OwnedUndo::RemoveNode => {
                    self.nodes.pop().unwrap();
                }
                OwnedUndo::RemoveValue => {
                    self.values.pop().unwrap();
                }
                OwnedUndo::RestoreValue(value) => self.values.push(value),
                OwnedUndo::SplitUnary { child, split } => {
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    self.nodes[child.0] = Some(split(parent));
                }
                OwnedUndo::SplitBinary { left, right, split } => {
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    let (a, b) = split(parent);
                    self.nodes[left.0] = Some(a);
                    self.nodes[right.0] = Some(b);
                }
            }
        }
    }
    /// Release old undo marks after the whole speculative parse commits.
    pub fn commit(&mut self) {
        self.undo.clear();
    }
    /// Finish ownership transfer without traversing or cloning the AST. This
    /// consumes the arena; old IDs/checkpoints cannot subsequently be used.
    pub fn into_root(mut self, root: Id) -> T {
        self.nodes[root.0].take().expect("consumed semantic root")
    }
}
