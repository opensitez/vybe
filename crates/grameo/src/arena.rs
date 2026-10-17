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
/// This alias keeps edit and composition metadata identical, preserving the
/// two-parameter API and inference for unannotated arena construction.
pub type MetadataArena<T, U> = EditArena<T, U, U>;

/// Opt-in arena with separate edit undo data `E` and composition metadata `U`.
/// Frequent scope/span edits need not reserve space for a large composition;
/// both payloads still use one transactional journal and checkpoint cursor.
pub struct EditArena<T, U, E> {
    nodes: Vec<Option<T>>,
    values: Vec<Id>,
    undo: Vec<OwnedUndo>,
    restored_values: Vec<Id>,
    edits: Vec<MetadataEdit<T, E>>,
    lists: Vec<ListComposition<T>>,
    metadata_lists: Vec<MetadataListComposition<T, U>>,
    binary_metadata: Vec<BinaryComposition<T, U>>,
    unary_metadata: Vec<UnaryMetadataComposition<T, U>>,
    unary_compositions: Vec<UnaryComposition<T>>,
    binary_compositions: Vec<PlainBinaryComposition<T>>,
}
struct ListComposition<T> {
    children: Vec<Id>,
    split: fn(T) -> Vec<T>,
}
struct MetadataListComposition<T, U> {
    children: Vec<Id>,
    previous: U,
    split: fn(T, U) -> Vec<T>,
}
struct MetadataEdit<T, U> {
    node: Id,
    previous: U,
    restore: fn(&mut T, U),
}
struct BinaryComposition<T, U> {
    left: Id,
    right: Id,
    previous: U,
    split: fn(T, U) -> (T, T),
}
struct UnaryComposition<T> {
    child: Id,
    split: fn(T) -> T,
}
struct UnaryMetadataComposition<T, U> {
    child: Id,
    previous: U,
    split: fn(T, U) -> T,
}
struct PlainBinaryComposition<T> {
    left: Id,
    right: Id,
    split: fn(T) -> (T, T),
}
// Keep the hot journal independent of the largest inverse operation. Push,
// pop and allocation are far more frequent than unary/binary compositions.
#[repr(u8)]
enum OwnedUndo {
    RemoveNode,
    RemoveValue,
    RestoreValue,
    RestorePenultimate,
    RestoreEdit,
    SplitList,
    SplitMetadataList,
    SplitBinaryMetadata,
    SplitUnary,
    SplitUnaryMetadata,
    SplitBinary,
}

#[cfg(test)]
mod journal_layout_tests {
    #[test]
    fn frequent_owned_undo_entries_fit_one_byte() {
        assert_eq!(std::mem::size_of::<super::OwnedUndo>(), 1);
    }

    #[test]
    fn wrapper_removal_restores_nested_speculation_and_metadata_edits() {
        let mut arena = super::EditArena::<String, (), String>::default();
        let prefix = arena.alloc("prefix".into());
        let outer = arena.alloc("outer".into());
        let inner = arena.alloc("inner".into());
        let child = arena.alloc("child".into());
        for id in [prefix, outer, inner, child] { arena.push(id); }
        let initial = arena.checkpoint();
        assert_eq!(arena.remove_value_below_top(), Some(inner));
        let nested = arena.checkpoint();
        assert_eq!(arena.remove_value_below_top(), Some(outer));
        arena.edit(child, |value| std::mem::replace(value, "edited".into()),
            |value, previous| *value = previous);
        let popped = arena.pop().unwrap();
        let parent = arena.unary(popped, |value| format!("({value})"),
            |value| value[1..value.len()-1].to_string());
        arena.push(parent);
        arena.rollback(nested);
        assert_eq!(arena.values(), &[prefix, outer, child]);
        assert_eq!(arena.get(child), "child");
        arena.rollback(initial);
        assert_eq!(arena.values(), &[prefix, outer, inner, child]);
        assert_eq!(arena.get(inner), "inner");
        assert_eq!(arena.get(outer), "outer");
    }

    #[test]
    fn wrapper_removal_matches_three_operation_route_and_commit() {
        for count in 0..32 {
            let mut arena = super::EditArena::<usize, (), ()>::default();
            let ids: Vec<_> = (0..count).map(|value| arena.alloc(value)).collect();
            for &id in &ids { arena.push(id); }
            let mark = arena.checkpoint();
            let removed = arena.remove_value_below_top();
            if count < 2 {
                assert!(removed.is_none());
                assert_eq!(arena.checkpoint(), mark);
                assert_eq!(arena.values(), ids);
            } else {
                assert_eq!(removed, Some(ids[count-2]));
                let mut expected = ids.clone(); expected.remove(count-2);
                assert_eq!(arena.values(), expected);
                assert_eq!(arena.checkpoint(), mark + 1);
            }
            arena.rollback(mark);
            assert_eq!(arena.values(), ids);
            if count >= 2 { arena.remove_value_below_top(); }
            let committed = arena.values().to_vec();
            arena.commit();
            assert_eq!(arena.checkpoint(), 0);
            assert_eq!(arena.values(), committed);
            for (value, id) in ids.iter().enumerate() { assert_eq!(arena.get(*id), &value); }
        }
    }
}
impl<T, U, E> Default for EditArena<T, U, E> {
    fn default() -> Self {
        Self {
            nodes: Vec::new(),
            values: Vec::new(),
            undo: Vec::new(),
            restored_values: Vec::new(),
            edits: Vec::new(),
            lists: Vec::new(),
            metadata_lists: Vec::new(),
            binary_metadata: Vec::new(),
            unary_metadata: Vec::new(),
            unary_compositions: Vec::new(),
            binary_compositions: Vec::new(),
        }
    }
}
impl<T, U, E> EditArena<T, U, E> {
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
    pub fn edit(&mut self, node: Id, change: impl FnOnce(&mut T) -> E, restore: fn(&mut T, E)) {
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
        self.restored_values.push(id);
        self.undo.push(OwnedUndo::RestoreValue);
        Some(id)
    }
    /// Remove the wrapper immediately below the top value while preserving
    /// the child and every earlier stack entry. Empty and singleton stacks
    /// are unchanged. Rollback restores the exact wrapper/child ordering.
    pub fn remove_value_below_top(&mut self) -> Option<Id> {
        let index = self.values.len().checked_sub(2)?;
        let removed = self.values[index];
        self.values[index] = self.values[index + 1];
        self.values.pop();
        self.restored_values.push(removed);
        self.undo.push(OwnedUndo::RestorePenultimate);
        Some(removed)
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
        self.unary_compositions.push(UnaryComposition { child, split });
        self.undo.push(OwnedUndo::SplitUnary);
        id
    }
    /// Compose an owned operand while journaling binding-selected inverse
    /// metadata. Prefix/postfix hooks can change spans or lowering state
    /// without cloning the operand or storing rollback data in the AST.
    /// `split` must recover the exact original child; constructors and inverses
    /// must not panic or mutate external state, as with `unary`.
    pub fn unary_with_metadata(
        &mut self, child: Id,
        make: impl FnOnce(T) -> (T, U), split: fn(T, U) -> T,
    ) -> Id {
        let child_value = self.nodes[child.0].take().expect("consumed semantic node");
        let (value, previous) = make(child_value);
        let id = Id(self.nodes.len());
        self.nodes.push(Some(value));
        self.unary_metadata.push(UnaryMetadataComposition { child, previous, split });
        self.undo.push(OwnedUndo::SplitUnaryMetadata);
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
        self.binary_compositions.push(PlainBinaryComposition { left, right, split });
        self.undo.push(OwnedUndo::SplitBinary);
        id
    }
    /// Combine owned children while retaining flat, binding-selected inverse
    /// metadata. Payloads live in the journal, rather than recursively inside
    /// the AST; rollback restores their original metadata without cloning.
    pub fn binary_with_metadata(
        &mut self, left: Id, right: Id,
        make: impl FnOnce(T, T) -> (T, U), split: fn(T, U) -> (T, T),
    ) -> Id {
        assert_ne!(left, right, "owned AST children cannot alias");
        assert!(self.nodes[left.0].is_some() && self.nodes[right.0].is_some(), "consumed semantic node");
        let left_value = self.nodes[left.0].take().unwrap();
        let right_value = self.nodes[right.0].take().unwrap();
        let (value, previous) = make(left_value, right_value);
        let id = Id(self.nodes.len());
        self.nodes.push(Some(value));
        self.binary_metadata.push(BinaryComposition { left, right, previous, split });
        self.undo.push(OwnedUndo::SplitBinaryMetadata);
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
        let values = self.take_children(&children);
        let id = Id(self.nodes.len());
        self.nodes.push(Some(make(values)));
        self.lists.push(ListComposition { children, split });
        self.undo.push(OwnedUndo::SplitList);
        id
    }
    /// Compose owned fields while preserving target-selected inverse metadata.
    /// This permits reversible lowering without retaining another AST tree.
    pub fn list_with_metadata(
        &mut self,
        children: Vec<Id>,
        make: impl FnOnce(Vec<T>) -> (T, U),
        split: fn(T, U) -> Vec<T>,
    ) -> Id {
        let values = self.take_children(&children);
        let (node, previous) = make(values);
        let id = Id(self.nodes.len());
        self.nodes.push(Some(node));
        self.metadata_lists.push(MetadataListComposition { children, previous, split });
        self.undo.push(OwnedUndo::SplitMetadataList);
        id
    }
    fn take_children(&mut self, children: &[Id]) -> Vec<T> {
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
            let mut unique = children.to_vec();
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
        children
            .iter()
            .map(|id| self.nodes[id.0].take().unwrap())
            .collect()
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
                OwnedUndo::RestoreValue => self.values.push(self.restored_values.pop().unwrap()),
                OwnedUndo::RestorePenultimate => {
                    let top = self.values.pop().unwrap();
                    self.values.push(self.restored_values.pop().unwrap());
                    self.values.push(top);
                }
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
                OwnedUndo::SplitBinaryMetadata => {
                    let BinaryComposition { left, right, previous, split } = self.binary_metadata.pop().unwrap();
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    let (a, b) = split(parent, previous);
                    self.nodes[left.0] = Some(a);
                    self.nodes[right.0] = Some(b);
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
                OwnedUndo::SplitMetadataList => {
                    let MetadataListComposition { children, previous, split } = self.metadata_lists.pop().unwrap();
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    let values = split(parent, previous);
                    assert_eq!(children.len(), values.len(), "list inverse changed child count");
                    for (child, value) in children.into_iter().zip(values) {
                        self.nodes[child.0] = Some(value);
                    }
                }
                OwnedUndo::SplitUnary => {
                    let UnaryComposition { child, split } = self.unary_compositions.pop().unwrap();
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    self.nodes[child.0] = Some(split(parent));
                }
                OwnedUndo::SplitUnaryMetadata => {
                    let UnaryMetadataComposition { child, previous, split } = self.unary_metadata.pop().unwrap();
                    let parent = self.nodes.pop().unwrap().expect("consumed parent");
                    self.nodes[child.0] = Some(split(parent, previous));
                }
                OwnedUndo::SplitBinary => {
                    let PlainBinaryComposition { left, right, split } = self.binary_compositions.pop().unwrap();
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
        self.restored_values.clear();
        self.edits.clear();
        self.lists.clear();
        self.metadata_lists.clear();
        self.binary_metadata.clear();
        self.unary_metadata.clear();
        self.unary_compositions.clear();
        self.binary_compositions.clear();
    }
    /// Commit multiple independent values without cloning. Consumed children
    /// are absent; live values retain ownership. All old marks and IDs expire.
    pub fn into_live_nodes(self) -> impl Iterator<Item = T> {
        self.nodes.into_iter().flatten()
    }
    /// Finish ownership transfer without traversing or cloning the AST. This
    /// consumes the arena; old IDs/checkpoints cannot subsequently be used.
    pub fn into_root(mut self, root: Id) -> T {
        self.nodes[root.0].take().expect("consumed semantic root")
    }
}
