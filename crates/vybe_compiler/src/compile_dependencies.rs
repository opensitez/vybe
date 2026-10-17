use std::collections::{HashMap, HashSet};
use vybe_runtime::{ExportEntry, ModuleRecord, Value};
#[derive(Clone, Debug, PartialEq, Eq)]
enum Terminal {
    Named(String, String),
    Scalar(Scalar),
}
#[derive(Clone, Debug, PartialEq, Eq)]
enum Scalar {
    Null,
    Undefined,
    Bool(bool),
    I32(i32),
    I64(i64),
    F32(u32),
    F64(u64),
    String(String),
    V128([u8; 16]),
}
fn scalar(value: &Value) -> Option<Scalar> {
    Some(match value {
        Value::Null => Scalar::Null,
        Value::Undefined => Scalar::Undefined,
        Value::Bool(v) => Scalar::Bool(*v),
        Value::I32(v) => Scalar::I32(*v),
        Value::I64(v) => Scalar::I64(*v),
        Value::F32(v) => Scalar::F32(v.to_bits()),
        Value::F64(v) => Scalar::F64(v.to_bits()),
        Value::String(v) => Scalar::String(v.to_string()),
        Value::V128(v) => Scalar::V128(*v),
        // Type-index values and object/symbol identity are not transferable scalar facts.
        _ => return None,
    })
}
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
fn scalar_matches(value: &Value, expected: &Scalar) -> bool {
    match (value, expected) {
        (Value::Null, Scalar::Null) | (Value::Undefined, Scalar::Undefined) => true,
        (Value::Bool(a), Scalar::Bool(b)) => a == b,
        (Value::I32(a), Scalar::I32(b)) => a == b,
        (Value::I64(a), Scalar::I64(b)) => a == b,
        (Value::F32(a), Scalar::F32(b)) => a.to_bits() == *b,
        (Value::F64(a), Scalar::F64(b)) => a.to_bits() == *b,
        (Value::String(a), Scalar::String(b)) => a.as_ref() == b,
        (Value::V128(a), Scalar::V128(b)) => a == b,
        _ => false,
    }
}
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum ReadKind {
    Export,
    Scalar,
}
/// Why an immutable speculative view cannot be reused. Names identify the
/// compiler read; diagnostics never expose the values stored in the VM.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ValidationFailure<'a> {
    Rejected,
    ExportChanged { module: &'a str, name: &'a str },
    ScalarChanged { module: &'a str, name: &'a str },
}
#[derive(Default, Debug)]
pub struct Certificate {
    // Borrow module/name keys on repeated reads instead of allocating tuple keys.
    reads: HashMap<String, [HashMap<String, Option<Terminal>>; 2]>,
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
        if self
            .read(ReadKind::Export, module, name)
            .is_some_and(|old| match (old, target) {
                (None, None) => true,
                (Some(Terminal::Named(a, b)), Some((m, n))) => a == m && b == n,
                _ => false,
            })
        {
            return;
        }
        self.record(
            ReadKind::Export,
            module,
            name,
            target.map(|(m, n)| Terminal::Named(m.clone(), n.clone())),
        );
    }
    pub fn record_scalar(&mut self, module: &str, name: &str, value: Option<&Value>) {
        if self
            .read(ReadKind::Scalar, module, name)
            .is_some_and(|old| match (old, value) {
                (None, None) => true,
                (Some(Terminal::Scalar(expected)), Some(value)) => scalar_matches(value, expected),
                _ => false,
            })
        {
            return;
        }
        let terminal = match value {
            None => None,
            Some(value) => match scalar(value) {
                Some(v) => Some(Terminal::Scalar(v)),
                None => {
                    self.rejected = true;
                    return;
                }
            },
        };
        self.record(ReadKind::Scalar, module, name, terminal);
    }
    fn read(&self, kind: ReadKind, module: &str, name: &str) -> Option<&Option<Terminal>> {
        self.reads.get(module)?.get(kind as usize)?.get(name)
    }
    fn record(&mut self, kind: ReadKind, module: &str, name: &str, terminal: Option<Terminal>) {
        let rows = if let Some(rows) = self.reads.get_mut(module) {
            rows
        } else {
            self.reads.entry(module.to_owned()).or_default()
        };
        let names = &mut rows[kind as usize];
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
        for (module, rows) in &self.reads {
            for (index, names) in rows.iter().enumerate() {
                let kind = if index == 0 {
                    ReadKind::Export
                } else {
                    ReadKind::Scalar
                };
                for (name, expected) in names {
                    let resolved = resolve(modules, module, name);
                    let matches = match expected {
                        Some(Terminal::Named(m, n)) => {
                            resolved.is_some_and(|(owner, terminal, _)| owner == m && terminal == n)
                        }
                        Some(Terminal::Scalar(v)) => {
                            resolved.is_some_and(|(_, _, entry)| match entry {
                                ExportEntry::Value(value) => scalar_matches(value, v),
                                _ => false,
                            })
                        }
                        None => match kind {
                            ReadKind::Export => resolved.is_none(),
                            ReadKind::Scalar => !resolved.is_some_and(|(_, _, entry)| {
                                matches!(entry, ExportEntry::Value(_))
                            }),
                        },
                    };
                    if !matches {
                        return Some(match kind {
                            ReadKind::Export => ValidationFailure::ExportChanged { module, name },
                            ReadKind::Scalar => ValidationFailure::ScalarChanged { module, name },
                        });
                    }
                }
            }
        }
        None
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn repeated_scalar_reads_still_replace_changed_and_negative_facts() {
        let mut modules = graph();
        let mut certificate = Certificate::default();
        for _ in 0..8 {
            certificate.record_scalar("a", "x", Some(&Value::F64(-0.0)));
        }
        assert!(certificate.validates(&modules));
        certificate.record_scalar("a", "x", Some(&Value::I32(7)));
        assert!(!certificate.validates(&modules));
        modules.get_mut("b").unwrap().exports.insert(
            "value".into(), ExportEntry::Value(Value::I32(7)),
        );
        assert!(certificate.validates(&modules));
        certificate.record_scalar("a", "x", None);
        assert!(!certificate.validates(&modules));
        modules.get_mut("b").unwrap().exports.insert(
            "value".into(), ExportEntry::Function { idx: 17 },
        );
        assert!(certificate.validates(&modules));
        // Unsupported VM-local values must reject even after a reusable fact.
        certificate.record_scalar("a", "x", Some(&Value::TypedNull(12)));
        certificate.record_scalar("a", "x", None);
        assert!(!certificate.validates(&modules));
    }
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
    fn scalar_facts_preserve_signed_zero_and_nan_bits() {
        let mut g = graph();
        let mut c = Certificate::default();
        c.record_scalar("a", "x", Some(&Value::F64(-0.0)));
        assert!(c.validates(&g));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Value(Value::F64(0.0)));
        assert!(!c.validates(&g));
        let nan = f64::from_bits(0x7ff8000000001234);
        c.record_scalar("a", "x", Some(&Value::F64(nan)));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Value(Value::F64(nan)));
        assert!(c.validates(&g));
    }
    #[test]
    fn vm_index_value_refuses_certification() {
        let g = graph();
        let mut c = Certificate::default();
        c.record_scalar("a", "x", Some(&Value::TypedNull(12)));
        assert!(!c.validates(&g));
    }
    #[test]
    fn scalar_missing_is_distinct_from_export_lookup() {
        let mut g = graph();
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Function { idx: 1 });
        let mut c = Certificate::default();
        c.record_export("a", "x", Some(&("b".into(), "value".into())));
        c.record_scalar("a", "x", None);
        assert!(c.validates(&g));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Value(Value::I32(7)));
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
        let mut c = Certificate::default();
        c.record_scalar("a", "x", Some(&Value::F64(-0.0)));
        assert_eq!(c.validation_failure(&g), None);
        assert!(c.validates(&g));
        g.get_mut("b")
            .unwrap()
            .exports
            .insert("value".into(), ExportEntry::Value(Value::F64(0.0)));
        assert_eq!(
            c.validation_failure(&g),
            Some(ValidationFailure::ScalarChanged {
                module: "a",
                name: "x"
            })
        );
        assert!(!c.validates(&g));
        c.reject();
        assert_eq!(c.validation_failure(&g), Some(ValidationFailure::Rejected));
        assert!(!c.validates(&g));
    }
}
