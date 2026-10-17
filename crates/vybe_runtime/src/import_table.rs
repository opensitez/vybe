//! Ordered import indices with mutation-observed interning.
//!
//! The vector is the guest ABI. A public mutable borrow invalidates the index
//! before yielding access, so renames, reorders and earlier duplicates retain
//! exactly the same first-occurrence behavior as an ordered linear search.
use crate::chunk::{FxBuildHasher, Import};
use std::{
    collections::HashMap,
    fmt,
    ops::{Deref, DerefMut},
    sync::Arc,
};

type Index = HashMap<String, HashMap<String, u16, FxBuildHasher>, FxBuildHasher>;

/// Validation-only selection. Neither the TLS nor its branch exists in normal
/// builds; the same guard is used for both lanes of a paired compiler sample.
#[cfg(feature = "primitive-benchmarks")]
#[derive(Clone, Copy)]
pub enum ImportInterningMode {
    Indexed,
    LinearControl,
}
#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static INTERNING_MODE: std::cell::Cell<ImportInterningMode> =
        const { std::cell::Cell::new(ImportInterningMode::Indexed) };
}
#[cfg(feature = "primitive-benchmarks")]
pub struct ImportInterningGuard {
    previous: ImportInterningMode,
    _same_thread: std::marker::PhantomData<std::rc::Rc<()>>,
}
#[cfg(feature = "primitive-benchmarks")]
impl ImportInterningGuard {
    pub fn enter(mode: ImportInterningMode) -> Self {
        Self {
            previous: INTERNING_MODE.with(|current| current.replace(mode)),
            _same_thread: std::marker::PhantomData,
        }
    }
}
#[cfg(feature = "primitive-benchmarks")]
impl Drop for ImportInterningGuard {
    fn drop(&mut self) {
        INTERNING_MODE.with(|current| current.set(self.previous));
    }
}

#[derive(Clone, Default)]
pub struct ImportTable {
    rows: Vec<Import>,
    index: Option<Arc<Index>>,
}

impl ImportTable {
    pub fn intern(&mut self, module: &str, name: &str) -> u16 {
        if self.index.is_none() {
            let mut index = Index::default();
            for (i, row) in self.rows.iter().enumerate() {
                index
                    .entry(row.module.clone())
                    .or_default()
                    .entry(row.name.clone())
                    .or_insert(i as u16);
            }
            self.index = Some(Arc::new(index));
        }
        let cached = self
            .index
            .as_ref()
            .unwrap()
            .get(module)
            .and_then(|names| names.get(name))
            .copied();
        #[cfg(feature = "primitive-benchmarks")]
        if INTERNING_MODE.with(|mode| matches!(mode.get(), ImportInterningMode::LinearControl)) {
            // Reproduce the old conservative prefix search even when the
            // cached row is valid. Only this search differs between lanes.
            let end = cached
                .filter(|i| {
                    self.rows
                        .get(*i as usize)
                        .is_some_and(|row| row.module == module && row.name == name)
                })
                .map_or(self.rows.len(), |i| i as usize + 1);
            if let Some(i) = self.rows[..end]
                .iter()
                .position(|row| row.module == module && row.name == name)
            {
                return i as u16;
            }
        } else if let Some(i) = cached {
            return i;
        }
        #[cfg(not(feature = "primitive-benchmarks"))]
        if let Some(i) = cached {
            return i;
        }
        let i = self.rows.len() as u16;
        self.rows.push(Import {
            module: module.into(),
            name: name.into(),
        });
        Arc::make_mut(self.index.as_mut().unwrap())
            .entry(module.into())
            .or_default()
            .insert(name.into(), i);
        i
    }

    pub fn into_vec(self) -> Vec<Import> {
        self.rows
    }

    #[cfg(test)]
    pub(crate) fn shares_index(&self, other: &Self) -> bool {
        match (&self.index, &other.index) {
            (Some(a), Some(b)) => Arc::ptr_eq(a, b),
            _ => false,
        }
    }
}

impl fmt::Debug for ImportTable {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.rows.fmt(f)
    }
}
impl PartialEq for ImportTable {
    fn eq(&self, rhs: &Self) -> bool {
        self.rows == rhs.rows
    }
}
impl Eq for ImportTable {}
impl PartialEq<Vec<Import>> for ImportTable {
    fn eq(&self, rhs: &Vec<Import>) -> bool {
        self.rows == *rhs
    }
}
impl PartialEq<ImportTable> for Vec<Import> {
    fn eq(&self, rhs: &ImportTable) -> bool {
        *self == rhs.rows
    }
}
impl From<Vec<Import>> for ImportTable {
    fn from(rows: Vec<Import>) -> Self {
        Self { rows, index: None }
    }
}
impl From<ImportTable> for Vec<Import> {
    fn from(table: ImportTable) -> Self {
        table.rows
    }
}
impl FromIterator<Import> for ImportTable {
    fn from_iter<T: IntoIterator<Item = Import>>(iter: T) -> Self {
        Vec::from_iter(iter).into()
    }
}
impl Deref for ImportTable {
    type Target = Vec<Import>;
    fn deref(&self) -> &Self::Target {
        &self.rows
    }
}
impl DerefMut for ImportTable {
    fn deref_mut(&mut self) -> &mut Self::Target {
        self.index = None;
        &mut self.rows
    }
}
impl AsRef<[Import]> for ImportTable {
    fn as_ref(&self) -> &[Import] {
        &self.rows
    }
}
impl IntoIterator for ImportTable {
    type Item = Import;
    type IntoIter = std::vec::IntoIter<Import>;
    fn into_iter(self) -> Self::IntoIter {
        self.rows.into_iter()
    }
}
impl<'a> IntoIterator for &'a ImportTable {
    type Item = &'a Import;
    type IntoIter = std::slice::Iter<'a, Import>;
    fn into_iter(self) -> Self::IntoIter {
        self.rows.iter()
    }
}
impl<'a> IntoIterator for &'a mut ImportTable {
    type Item = &'a mut Import;
    type IntoIter = std::slice::IterMut<'a, Import>;
    fn into_iter(self) -> Self::IntoIter {
        self.index = None;
        self.rows.iter_mut()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn oracle(rows: &mut Vec<Import>, module: &str, name: &str) -> u16 {
        if let Some(i) = rows
            .iter()
            .position(|r| r.module == module && r.name == name)
        {
            i as u16
        } else {
            let i = rows.len() as u16;
            rows.push(Import {
                module: module.into(),
                name: name.into(),
            });
            i
        }
    }
    #[test]
    fn arbitrary_mutations_preserve_first_occurrence() {
        let mut actual = ImportTable::default();
        let mut expected: Vec<Import> = Vec::new();
        let mut seed = 72458211u64;
        for step in 0..20_000 {
            seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
            let module = format!("module{}", (seed >> 24) % 3);
            let name = format!("name{}", (seed >> 32) % 23);
            match (seed >> 48) % 9 {
                0 if !expected.is_empty() => {
                    let i = seed as usize % expected.len();
                    actual[i].name = name.clone();
                    expected[i].name = name.clone();
                }
                1 => {
                    actual.reverse();
                    expected.reverse();
                }
                2 if !expected.is_empty() => {
                    let i = seed as usize % expected.len();
                    actual.remove(i);
                    expected.remove(i);
                }
                3 if step % 17 == 0 => {
                    actual.clear();
                    expected.clear();
                }
                4 => {
                    let row = Import {
                        module: module.clone(),
                        name: name.clone(),
                    };
                    actual.push(row.clone());
                    expected.push(row);
                }
                _ => {}
            }
            assert_eq!(
                actual.intern(&module, &name),
                oracle(&mut expected, &module, &name),
                "step {step}"
            );
            assert_eq!(actual, expected);
        }
    }
    #[test]
    fn cloning_shares_index_until_a_mutation_and_debug_is_the_vector() {
        let mut a = ImportTable::default();
        a.intern("m", "a");
        let mut b = a.clone();
        assert!(a.shares_index(&b));
        b[0].name = "b".into();
        assert_eq!(a.intern("m", "a"), 0);
        assert_eq!(b.intern("m", "a"), 1);
        let rows = vec![
            Import {
                module: "x".into(),
                name: "y".into()
            };
            2
        ];
        let mut a: ImportTable = rows.clone().into();
        assert_eq!(a.intern("x", "y"), 0);
        assert_eq!(format!("{a:?}"), format!("{rows:?}"));
    }
    #[test]
    fn mutable_iteration_invalidates_and_replacement_starts_dirty() {
        let mut a = ImportTable::default();
        a.intern("m", "a");
        for row in &mut a {
            row.name = "b".into();
        }
        assert_eq!(a.intern("m", "a"), 1);
        a = vec![Import {
            module: "m".into(),
            name: "b".into(),
        }]
        .into();
        assert_eq!(a.intern("m", "b"), 0);
        assert_eq!(a.intern("m", "a"), 1);
    }

    #[test]
    fn module_rename_swap_extend_and_u16_indices_match_the_vector() {
        let mut a = ImportTable::default();
        a.intern("a", "name"); a.intern("b", "name");
        a[0].module = "b".into();
        assert_eq!(a.intern("b", "name"), 0);
        a.swap(0, 1);
        a.extend([Import { module: "c".into(), name: "extra".into() }]);
        assert_eq!(a.intern("c", "extra"), 2);
        let rows: Vec<_> = (0..65_538).map(|i| Import {
            module: "wide".into(), name: i.to_string(),
        }).collect();
        let mut a: ImportTable = rows.into();
        assert_eq!(a.intern("wide", "65535"), u16::MAX);
        assert_eq!(a.intern("wide", "65536"), 0);
        a[0].name = "65536".into();
        assert_eq!(a.intern("wide", "65536"), 0);
        assert_eq!(a.intern("wide", "65537"), 1);
    }

    #[cfg(feature = "primitive-benchmarks")]
    #[test]
    fn validation_scope_is_nested_thread_local_and_restores_on_unwind() {
        use ImportInterningMode::*;
        assert!(INTERNING_MODE.with(|m| matches!(m.get(), Indexed)));
        let _outer = ImportInterningGuard::enter(LinearControl);
        { let _inner = ImportInterningGuard::enter(Indexed);
          assert!(INTERNING_MODE.with(|m| matches!(m.get(), Indexed))); }
        assert!(INTERNING_MODE.with(|m| matches!(m.get(), LinearControl)));
        std::thread::spawn(|| assert!(INTERNING_MODE.with(|m| matches!(m.get(), Indexed))))
            .join().unwrap();
        let _ = std::panic::catch_unwind(|| {
            let _inner = ImportInterningGuard::enter(Indexed); panic!("scope restore");
        });
        assert!(INTERNING_MODE.with(|m| matches!(m.get(), LinearControl)));
    }
}
