use std::collections::HashSet;
/// Keep exact declarations authoritative; the optional ASCII index only
/// accelerates queries which already used `eq_ignore_ascii_case`.
/// No mutable set access is exposed, so insertions cannot bypass the index.
pub(crate) type BodySymbols = IndexedSymbols<std::collections::hash_map::RandomState>;
pub(crate) type FunctionSymbols = IndexedSymbols<vybe_runtime::chunk::FxBuildHasher>;

#[derive(Debug, Clone, Default)]
pub(crate) struct IndexedSymbols<S: std::hash::BuildHasher> {
    exact: HashSet<String, S>,
    folded: Option<HashSet<String>>,
}
impl<S: std::hash::BuildHasher + Default> IndexedSymbols<S> {
    pub(crate) fn new() -> Self {
        Self::default()
    }
    pub(crate) fn configure_folded(&mut self, enabled: bool) {
        if enabled {
            if self.folded.is_none() {
                self.folded = Some(self.exact.iter().map(|v| v.to_ascii_lowercase()).collect());
            }
        } else {
            self.folded = None;
        }
    }
    pub(crate) fn folded_index_enabled(&self) -> bool {
        self.folded.is_some()
    }
    pub(crate) fn insert(&mut self, value: String) -> bool {
        if let Some(f) = &mut self.folded {
            f.insert(value.to_ascii_lowercase());
        }
        self.exact.insert(value)
    }
    pub(crate) fn contains_ascii_folded(&self, value: &str) -> bool {
        if self.exact.contains(value) {
            return true;
        }
        match &self.folded {
            None => self.exact.iter().any(|n| n.eq_ignore_ascii_case(value)),
            Some(f) => {
                if value.bytes().any(|b| b.is_ascii_uppercase()) {
                    f.contains(value.to_ascii_lowercase().as_str())
                } else {
                    f.contains(value)
                }
            }
        }
    }
}
impl<S: std::hash::BuildHasher> std::ops::Deref for IndexedSymbols<S> {
    type Target = HashSet<String, S>;
    fn deref(&self) -> &Self::Target {
        &self.exact
    }
}
impl<'a, S: std::hash::BuildHasher> IntoIterator for &'a IndexedSymbols<S> {
    type Item = &'a String;
    type IntoIter = std::collections::hash_set::Iter<'a, String>;
    fn into_iter(self) -> Self::IntoIter {
        self.exact.iter()
    }
}

#[cfg(test)]
mod tests {
    use super::BodySymbols;

    #[test]
    fn folded_index_preserves_exact_ascii_and_non_ascii_membership() {
        let names = [
            "Alpha", "aLPHa", "ÄName", "äname", "same", "N.Type", "n.type",
        ];
        let mut symbols = BodySymbols::new();
        for name in names {
            symbols.insert(name.into());
        }
        for enabled in [false, true, true, false, true] {
            symbols.configure_folded(enabled);
            for needle in [
                "alpha", "ALPHA", "ÄName", "äname", "ÄNAME", "n.TYPE", "missing",
            ] {
                assert_eq!(
                    symbols.contains_ascii_folded(needle),
                    names.iter().any(|name| name.eq_ignore_ascii_case(needle))
                );
            }
        }
        symbols.insert("NewName".into());
        assert!(symbols.contains_ascii_folded("newname"));
        assert!(!symbols.contains("newname"));
        let mut cloned = symbols.clone();
        cloned.insert("OnlyClone".into());
        assert!(!symbols.contains_ascii_folded("onlyclone"));
        cloned.configure_folded(false);
        assert!(cloned.contains_ascii_folded("onlyclone"));
    }
}

#[cfg(test)]
#[test]
fn function_index_keeps_fx_exact_set_and_original_iteration() {
    let names = ["Function", "function", "ÄName", "äname", "Case"];
    let mut indexed = FunctionSymbols::new();
    let mut control =
        std::collections::HashSet::<String, vybe_runtime::chunk::FxBuildHasher>::default();
    for name in names {
        assert_eq!(indexed.insert(name.into()), control.insert(name.into()));
    }
    for enabled in [false, true, false, true] {
        indexed.configure_folded(enabled);
        assert_eq!(
            indexed.iter().collect::<Vec<_>>(),
            control.iter().collect::<Vec<_>>()
        );
        for needle in ["Function", "FUNCTION", "ÄNAME", "äname", "missing"] {
            assert_eq!(indexed.contains(needle), control.contains(needle));
            assert_eq!(
                indexed.contains_ascii_folded(needle),
                control.iter().any(|name| name.eq_ignore_ascii_case(needle))
            );
        }
    }
    let mut cloned = indexed.clone();
    cloned.insert("LateFunction".into());
    assert!(cloned.contains_ascii_folded("latefunction"));
    assert!(!indexed.contains_ascii_folded("latefunction"));
}
