/// Cache exact segmented queries independently for each folding policy.
/// Bounds only limit cache storage; deeper queries still use the resolver.
use super::namespaces::{Fold, ResolutionTarget};
use std::collections::HashMap;
#[derive(Default)]
struct Node {
    children: HashMap<String, Node>,
    values: [Option<Option<ResolutionTarget>>; 3],
}
#[derive(Default)]
pub(crate) struct UserTreeCache {
    enabled: bool,
    root: Node,
    nodes: usize,
}
fn policy(fold: Fold) -> usize {
    match fold {
        None => 0,
        Some(vybe_ast::CaseAlphabet::Ascii) => 1,
        Some(vybe_ast::CaseAlphabet::Unicode) => 2,
    }
}
impl UserTreeCache {
    pub(crate) fn configure(&mut self, enabled: bool) {
        if self.enabled != enabled {
            self.root = Node::default();
            self.nodes = 0;
        }
        self.enabled = enabled;
    }
    pub(crate) fn enabled(&self) -> bool {
        self.enabled
    }
    pub(crate) fn invalidate(&mut self) {
        if self.enabled {
            self.root = Node::default();
            self.nodes = 0;
        }
    }
    pub(crate) fn get(&self, path: &[&str], fold: Fold) -> Option<Option<ResolutionTarget>> {
        if !self.enabled || path.len() > 64 {
            return None;
        }
        let mut current = &self.root;
        for segment in path {
            current = current.children.get(*segment)?;
        }
        current.values[policy(fold)].clone()
    }
    pub(crate) fn insert(&mut self, path: &[&str], fold: Fold, value: Option<ResolutionTarget>) {
        if !self.enabled || path.len() > 64 {
            return;
        }
        if self.nodes + path.len() > 4096 {
            self.root = Node::default();
            self.nodes = 0;
        }
        let mut current = &mut self.root;
        for segment in path {
            current = current
                .children
                .entry((*segment).to_owned())
                .or_insert_with(|| {
                    self.nodes += 1;
                    Node::default()
                });
        }
        current.values[policy(fold)] = Some(value);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn segments_policy_negative_results_and_invalidation_remain_distinct() {
        let mut cache = UserTreeCache::default();
        let path = ["A.B", "C"];
        cache.insert(&path, None, None);
        assert!(cache.get(&path, None).is_none());
        cache.configure(true);
        cache.insert(&path, None, None);
        assert!(matches!(cache.get(&path, None), Some(None)));
        assert!(cache.get(&["A", "B.C"], None).is_none());
        let ascii = Some(vybe_ast::CaseAlphabet::Ascii);
        let unicode = Some(vybe_ast::CaseAlphabet::Unicode);
        assert!(cache.get(&path, ascii).is_none());
        cache.insert(&path, ascii, Some(ResolutionTarget::CommonEmit("x".into())));
        assert!(
            matches!(cache.get(&path, ascii), Some(Some(ResolutionTarget::CommonEmit(value))) if value == "x")
        );
        assert!(cache.get(&path, unicode).is_none());
        cache.invalidate();
        assert!(cache.get(&path, None).is_none());
        cache.insert(&path, None, None);
        cache.configure(false);
        cache.configure(true);
        assert!(cache.get(&path, None).is_none());
    }

    #[test]
    fn cache_bounds_evict_without_restricting_resolver_inputs() {
        let mut cache = UserTreeCache::default();
        cache.configure(true);
        let deep = vec!["x"; 65];
        cache.insert(&deep, None, None);
        assert!(cache.get(&deep, None).is_none());
        for number in 0..4100 {
            let segment = format!("node{number}");
            cache.insert(&[&segment], None, None);
        }
        assert!(cache.get(&["node0"], None).is_none());
        assert!(matches!(cache.get(&["node4099"], None), Some(None)));
    }
}
