use std::collections::{HashMap, HashSet};
use vybe_runtime::{ExportEntry, ModuleRecord};
fn resolve<'a>(
    modules: &'a HashMap<String, ModuleRecord>,
    module: &'a str,
    name: &'a str,
) -> Option<(&'a str, &'a str, &'a ExportEntry)> {
    let entry = modules.get(module)?.exports.get(name)?;
    if !matches!(entry, ExportEntry::Indirect { .. }) {
        return Some((module, name, entry));
    }
    let (mut module, mut name) = (module, name);
    let mut seen = HashSet::new();
    loop {
        if !seen.insert((module, name)) {
            return None;
        }
        let entry = modules.get(module)?.exports.get(name)?;
        match entry {
            ExportEntry::Indirect { from, name: target } => {
                module = from;
                name = target;
            }
            _ => return Some((module, name, entry)),
        }
    }
}
/// Why an immutable speculative view cannot be reused. Names identify the
/// compiler read; diagnostics never expose the values stored in the VM.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ValidationFailure<'a> {
    Rejected,
    ExportChanged { module: &'a str, name: &'a str },
}
#[derive(Default, Debug)]
pub struct Certificate {
    // Borrow module/name keys on repeated reads instead of allocating tuple keys.
    reads: HashMap<String, HashMap<String, Option<(String, String)>>>,
    rejected: bool,
    speculative: bool,
    diagnostics: Vec<String>,
    runtime_helpers: Vec<vybe_runtime::shared_provider::InstalledRuntimeProvider>,
}
impl Certificate {
    pub(crate) fn install_runtime_helpers(&mut self,
        providers: Vec<vybe_runtime::shared_provider::InstalledRuntimeProvider>) {
        self.runtime_helpers = providers;
    }
    pub(crate) fn runtime_helper_exports(&self, chunks: &[vybe_runtime::Chunk]) -> Vec<&'static str> {
        crate::primitives::bundle::installed_runtime_helper_exports(&self.runtime_helpers, chunks)
    }
    pub(crate) fn runtime_helpers_valid(&self, vm: &vybe_runtime::VM) -> bool {
        self.runtime_helpers.iter().all(|provider| vm.validate_runtime_helper(provider)
            && crate::primitives::bundle::runtime_provider_diagnostics_clean(vm, provider))
    }
    pub fn speculative() -> Self {
        Self {
            speculative: true,
            ..Self::default()
        }
    }
    pub fn capture_diagnostic(&mut self, message: &str) -> bool {
        if !self.speculative {
            return false;
        }
        self.diagnostics.push(message.into());
        true
    }
    pub fn take_diagnostics(&mut self) -> Vec<String> {
        std::mem::take(&mut self.diagnostics)
    }
    pub fn reject(&mut self) {
        self.rejected = true;
    }
    // Integrating lookup helpers call this only at actual compiler reads. No broad VM hash.
    pub fn record_export(&mut self, module: &str, name: &str, target: Option<&(String, String)>) {
        if self.read(module, name).is_some_and(|old| old.as_ref() == target) {
            return;
        }
        self.record(module, name, target.cloned());
    }
    fn read(&self, module: &str, name: &str) -> Option<&Option<(String, String)>> {
        self.reads.get(module)?.get(name)
    }
    fn record(&mut self, module: &str, name: &str, terminal: Option<(String, String)>) {
        let names = if let Some(names) = self.reads.get_mut(module) {
            names
        } else {
            self.reads.entry(module.to_owned()).or_default()
        };
        if let Some(old) = names.get_mut(name) {
            *old = terminal;
        } else {
            names.insert(name.to_owned(), terminal);
        }
    }
    pub fn validates(&self, modules: &HashMap<String, ModuleRecord>) -> bool {
        self.validation_failure(modules).is_none()
    }
    /// Return one failed read using exactly the normal acceptance predicates.
    /// When several reads changed, HashMap iteration does not define which is
    /// reported first. A rejected certificate does not inspect the module graph.
    pub fn validation_failure(
        &self,
        modules: &HashMap<String, ModuleRecord>,
    ) -> Option<ValidationFailure<'_>> {
        if self.rejected {
            return Some(ValidationFailure::Rejected);
        }
        for (module, names) in &self.reads {
            for (name, expected) in names {
                let resolved = resolve(modules, module, name);
                let matches = match expected {
                    Some((m, n)) => {
                        resolved.is_some_and(|(owner, terminal, _)| owner == m && terminal == n)
                    }
                    None => resolved.is_none(),
                };
                if !matches {
                    return Some(ValidationFailure::ExportChanged { module, name });
                }
            }
        }
        None
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    use vybe_runtime::Value;
    fn graph() -> HashMap<String, ModuleRecord> {
        let mut a = ModuleRecord::new_synthetic("a");
        a.exports.insert(
            "x".into(),
            ExportEntry::Indirect {
                from: "b".into(),
                name: "value".into(),
            },
        );
        let mut b = ModuleRecord::new_synthetic("b");
        b.exports
            .insert("value".into(), ExportEntry::Value(Value::F64(-0.0)));
        HashMap::from([("a".into(), a), ("b".into(), b)])
    }
    #[test]
    fn changed_selected_alias_invalidates_unrelated_export_does_not() {
        let mut g = graph();
        let mut c = Certificate::default();
        c.record_export("a", "x", Some(&("b".into(), "value".into())));
        assert!(c.validates(&g));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("unrelated".into(), ExportEntry::Function { idx: 999 });
        assert!(c.validates(&g));
        g.get_mut("a").unwrap().exports.insert(
            "x".into(),
            ExportEntry::Indirect {
                from: "b".into(),
                name: "unrelated".into(),
            },
        );
        assert!(!c.validates(&g));
    }
    #[test]
    fn named_lookup_matches_flattening_across_terminal_kind_changes() {
        let mut g = graph();
        let mut c = Certificate::default();
        c.record_export("a", "x", Some(&("b".into(), "value".into())));
        assert!(c.validates(&g));
        for entry in [
            ExportEntry::Function { idx: 17 },
            ExportEntry::ResourceType { type_id: 11 },
            ExportEntry::Class { type_id: 31 },
            ExportEntry::Value(Value::I32(9)),
        ] {
            g.get_mut("b")
                .unwrap()
                .exports
                .insert("value".into(), entry);
            assert!(c.validates(&g));
            assert_eq!(
                crate::bundle::flatten_module_exports(&g)["a"]["x"],
                ("b".into(), "value".into())
            );
        }
    }
    #[test]
    fn negative_lookup_becoming_present_invalidates() {
        let mut g = graph();
        let mut c = Certificate::default();
        c.record_export("a", "missing", None);
        assert!(c.validates(&g));
        g.get_mut("a")
            .unwrap()
            .exports
            .insert("missing".into(), ExportEntry::Function { idx: 1 });
        assert!(!c.validates(&g));
    }
    #[test]
    fn cyclic_alias_negative_lookup_is_checked() {
        let mut g = graph();
        g.get_mut("b").unwrap().exports.insert(
            "value".into(),
            ExportEntry::Indirect {
                from: "a".into(),
                name: "x".into(),
            },
        );
        let mut c = Certificate::default();
        c.record_export("a", "x", None);
        assert!(c.validates(&g));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Function { idx: 1 });
        assert!(!c.validates(&g));
    }
    #[test]
    fn failure_diagnostics_preserve_acceptance_and_read_names() {
        let mut g = graph();
        let mut c = Certificate::default();
        c.record_export("a", "missing", None);
        assert_eq!(c.validation_failure(&g), None);
        assert!(c.validates(&g));
        g.get_mut("a")
            .unwrap()
            .exports
            .insert("missing".into(), ExportEntry::Function { idx: 1 });
        assert_eq!(
            c.validation_failure(&g),
            Some(ValidationFailure::ExportChanged {
                module: "a",
                name: "missing"
            })
        );
        assert!(!c.validates(&g));
        c.reject();
        assert_eq!(c.validation_failure(&g), Some(ValidationFailure::Rejected));
        assert!(!c.validates(&g));
    }
}
