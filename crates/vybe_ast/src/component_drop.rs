//! Release owned nested component graphs without recursive Arc destruction.
use super::{ComponentBinary, ComponentBinaryDef};
use std::sync::Arc;

impl Drop for ComponentBinary {
    fn drop(&mut self) {
        let mut pending = std::mem::take(&mut self.definitions);
        while let Some(definition) = pending.pop() {
            if let ComponentBinaryDef::NestedComponent(child) = definition
                && let Some(mut component) = Arc::into_inner(child)
            {
                pending.append(&mut component.definitions);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn chain(depth: usize, leaf: Arc<ComponentBinary>) -> Arc<ComponentBinary> {
        (0..depth).fold(leaf, |child, _| Arc::new(ComponentBinary {
            definitions: vec![ComponentBinaryDef::NestedComponent(child)],
        }))
    }

    #[test]
    fn deep_owned_components_drop_on_a_small_stack() {
        std::thread::Builder::new().stack_size(64 * 1024).spawn(|| {
            let leaf = Arc::new(ComponentBinary::default());
            let weak = Arc::downgrade(&leaf);
            drop(chain(32_768, leaf));
            assert!(weak.upgrade().is_none());
        }).unwrap().join().unwrap();
    }

    #[test]
    fn shared_children_survive_until_their_last_owner() {
        let child = chain(32_768, Arc::new(ComponentBinary::default()));
        let weak = Arc::downgrade(&child);
        let root = ComponentBinary { definitions: vec![
            ComponentBinaryDef::NestedComponent(child.clone()),
            ComponentBinaryDef::NestedComponent(child.clone()),
        ] };
        drop(root);
        assert_eq!(Arc::strong_count(&child), 1);
        assert!(weak.upgrade().is_some());
        drop(child);
        assert!(weak.upgrade().is_none());
    }

    #[test]
    fn concurrent_last_owners_release_the_shared_graph() {
        let leaf = Arc::new(ComponentBinary::default());
        let weak = Arc::downgrade(&leaf);
        let root = chain(32_768, leaf);
        let barrier = Arc::new(std::sync::Barrier::new(2));
        let mut threads = Vec::new();
        for root in [root.clone(), root] {
            let barrier = barrier.clone();
            threads.push(std::thread::Builder::new().stack_size(64 * 1024)
                .spawn(move || { barrier.wait(); drop(root); }).unwrap());
        }
        for thread in threads { thread.join().unwrap(); }
        assert!(weak.upgrade().is_none());
    }

    #[test]
    fn clone_and_make_mut_keep_independent_parent_definitions() {
        let mut root = chain(2, Arc::new(ComponentBinary::default()));
        let copy = root.clone();
        Arc::make_mut(&mut root).definitions.push(ComponentBinaryDef::Custom {
            name: "retained".into(), data: vec![1, 2, 3],
        });
        assert_eq!(root.definitions.len(), 2);
        assert_eq!(copy.definitions.len(), 1);
    }
}
