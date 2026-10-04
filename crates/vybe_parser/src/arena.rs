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
pub type OwnedArena<T> = MetadataArena<T, ()>;

/// Owned AST arena with a binding-selected metadata undo payload. Choose a
/// small payload or enum for edits; ordinary composition uses `OwnedArena<T>`.
pub struct MetadataArena<T, U> {
    nodes: Vec<Option<T>>,
    values: Vec<Id>,
    undo: Vec<OwnedUndo<T>>,
    edits: Vec<MetadataEdit<T, U>>,
    lists: Vec<ListComposition<T>>,
}
struct ListComposition<T> {
    children: Vec<Id>,
    split: fn(T) -> Vec<T>,
}
struct MetadataEdit<T, U> {
    node: Id,
    previous: U,
    restore: fn(&mut T, U),
}
enum OwnedUndo<T> {
    RemoveNode,
    RemoveValue,
    RestoreValue(Id),
    RestoreEdit,
    SplitList,
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
impl<T, U> Default for MetadataArena<T, U> {
    fn default() -> Self {
        Self {
            nodes: Vec::new(),
            values: Vec::new(),
            undo: Vec::new(),
            edits: Vec::new(),
            lists: Vec::new(),
        }
    }
}
impl<T, U> MetadataArena<T, U> {
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
    /// Edit a live node while retaining only the metadata needed to undo it.
    /// `change` returns the previous metadata; `restore` must restore the exact
    /// node. Both functions must not panic or mutate external state. A typed
    /// payload (such as a span or a binding-owned enum) avoids AST cloning and
    /// boxed undo closures. Later compositions are undone before this edit.
    pub fn edit(&mut self, node: Id, change: impl FnOnce(&mut T) -> U, restore: fn(&mut T, U)) {
        let value = self.nodes[node.0].as_mut().expect("consumed semantic node");
        let previous = change(value);
        self.edits.push(MetadataEdit {
            node,
            previous,
            restore,
        });
        self.undo.push(OwnedUndo::RestoreEdit);
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
    /// Move an ordered list of exclusive children into a parent. `split` must
    /// recover exactly those children, in the same order. This supports module
    /// bodies, argument lists and collection fields without cloning or a final
    /// materialization walk. Empty lists are valid. Constructors/inverses have
    /// the same no-panic/no-external-effects contract as unary composition.
    pub fn list(
        &mut self,
        children: Vec<Id>,
        make: impl FnOnce(Vec<T>) -> T,
        split: fn(T) -> Vec<T>,
    ) -> Id {
        // Validate before taking any child. Small lists avoid an extra
        // allocation; large lists avoid quadratic duplicate checking.
        if children.len() <= 16 {
            for (index, child) in children.iter().enumerate() {
                assert!(
                    !children[..index].contains(child),
                    "owned AST children cannot alias"
                );
            }
        } else {
            let mut unique = children.clone();
            unique.sort_unstable_by_key(|id| id.0);
            assert!(
                unique.windows(2).all(|pair| pair[0] != pair[1]),
                "owned AST children cannot alias"
            );
        }
        assert!(
            children
                .iter()
                .all(|id| self.nodes.get(id.0).is_some_and(Option::is_some)),
            "invalid or consumed semantic node"
        );
        let values = children
            .iter()
            .map(|id| self.nodes[id.0].take().unwrap())
            .collect();
        let id = Id(self.nodes.len());
        self.nodes.push(Some(make(values)));
        self.lists.push(ListComposition { children, split });
        self.undo.push(OwnedUndo::SplitList);
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
                OwnedUndo::RestoreEdit => {
                    let MetadataEdit {
                        node,
                        previous,
                        restore,
                    } = self.edits.pop().unwrap();
                    restore(
                        self.nodes[node.0].as_mut().expect("consumed edited node"),
                        previous,
                    );
                }
                OwnedUndo::SplitList => {
                    let ListComposition { children, split } = self.lists.pop().unwrap();
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    let values = split(parent);
                    assert_eq!(
                        children.len(),
                        values.len(),
                        "list inverse changed child count"
                    );
                    for (child, value) in children.into_iter().zip(values) {
                        self.nodes[child.0] = Some(value);
                    }
                }
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
        self.edits.clear();
        self.lists.clear();
    }
    /// Finish ownership transfer without traversing or cloning the AST. This
    /// consumes the arena; old IDs/checkpoints cannot subsequently be used.
    pub fn into_root(mut self, root: Id) -> T {
        self.nodes[root.0].take().expect("consumed semantic root")
    }
}
