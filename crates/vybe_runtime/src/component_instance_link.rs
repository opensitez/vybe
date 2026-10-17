use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use vybe_ast::canon::ComponentItemKind;

use crate::canon_def::ComponentFuncLink;
use crate::VMError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InstanceExport {
    pub name: Arc<str>,
    pub kind: ComponentItemKind,
    pub index: u32,
    pub function_type: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InstanceExports {
    exports: Arc<[InstanceExport]>,
}

impl InstanceExports {
    pub fn new(exports: Vec<InstanceExport>) -> Result<Self, VMError> {
        let mut names = HashSet::with_capacity(exports.len());
        for export in &exports {
            if !names.insert(export.name.as_ref()) {
                return Err(VMError::new(format!(
                    "component instance: duplicate export `{}`",
                    export.name
                )));
            }
            if export.kind == ComponentItemKind::Func && export.function_type.is_none() {
                return Err(VMError::new(format!(
                    "component instance: function export `{}` has no type",
                    export.name
                )));
            }
        }
        Ok(Self {
            exports: exports.into(),
        })
    }

    pub fn exports(&self) -> &[InstanceExport] {
        &self.exports
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InstanceFunctionImport {
    pub instance: u32,
    pub name: Arc<str>,
    pub import: u32,
    pub expected_type: u32,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct InstanceLinkMetadata {
    pub instances: Vec<InstanceExports>,
    pub imports: Vec<InstanceFunctionImport>,
}

pub fn bind_instance_function_imports<'a>(
    existing: &'a [ComponentFuncLink],
    metadata: Option<&InstanceLinkMetadata>,
) -> Result<Cow<'a, [ComponentFuncLink]>, VMError> {
    match metadata {
        Some(metadata) if !metadata.imports.is_empty() => Ok(Cow::Owned(collect_with_prefix(
            existing,
            &metadata.instances,
            &metadata.imports,
        )?)),
        _ => Ok(Cow::Borrowed(existing)),
    }
}

/// Resolve named instance function exports once at instantiation. The existing
/// canonical function linker validates indices, signatures and binding cycles.
pub fn collect_instance_function_links(
    instances: &[InstanceExports],
    imports: &[InstanceFunctionImport],
) -> Result<Vec<ComponentFuncLink>, VMError> {
    collect_with_prefix(&[], instances, imports)
}

fn collect_with_prefix(
    existing: &[ComponentFuncLink],
    instances: &[InstanceExports],
    imports: &[InstanceFunctionImport],
) -> Result<Vec<ComponentFuncLink>, VMError> {
    const INDEX_THRESHOLD: usize = 8;
    let mut indexes: HashMap<u32, HashMap<&str, &InstanceExport>> = HashMap::new();
    let capacity = existing
        .len()
        .checked_add(imports.len())
        .ok_or_else(|| VMError::new("component link: binding count overflow"))?;
    let mut links = Vec::with_capacity(capacity);
    links.extend_from_slice(existing);
    let mut start = 0;
    while start < imports.len() {
        let instance_index = imports[start].instance;
        let instance = instances.get(instance_index as usize).ok_or_else(|| {
            VMError::new(format!(
                "component link: instance {instance_index} is not defined"
            ))
        })?;
        let index = if instance.exports.len() <= INDEX_THRESHOLD {
            None
        } else {
            Some(indexes.entry(instance_index).or_insert_with(|| {
                instance
                    .exports
                    .iter()
                    .map(|export| (export.name.as_ref(), export))
                    .collect()
            }))
        };
        while start < imports.len() && imports[start].instance == instance_index {
            let import = &imports[start];
            let export = match &index {
                Some(index) => index.get(import.name.as_ref()).copied(),
                None => instance
                    .exports
                    .iter()
                    .find(|export| export.name == import.name),
            }
            .ok_or_else(|| {
                VMError::new(format!(
                    "component link: instance {} has no export `{}`",
                    import.instance, import.name
                ))
            })?;
            if export.kind != ComponentItemKind::Func {
                return Err(VMError::new(format!(
                    "component link: instance export `{}` is {:?}, not a function",
                    import.name, export.kind
                )));
            }
            let supplied_type = export.function_type.ok_or_else(|| {
                VMError::new(format!(
                    "component link: function export `{}` has no type",
                    import.name
                ))
            })?;
            links.push(ComponentFuncLink {
                import: import.import,
                supplied: export.index,
                expected_type: import.expected_type,
                supplied_type,
            });
            start += 1;
        }
    }
    Ok(links)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn export(name: &str, index: u32) -> InstanceExport {
        InstanceExport {
            name: Arc::from(name),
            kind: ComponentItemKind::Func,
            index,
            function_type: Some(index + 1),
        }
    }

    fn import(name: &str, slot: u32) -> InstanceFunctionImport {
        InstanceFunctionImport {
            instance: 0,
            name: Arc::from(name),
            import: slot,
            expected_type: 1,
        }
    }

    #[test]
    fn small_and_indexed_exports_produce_the_same_canonical_links() {
        for count in [1, 8, 9, 128] {
            let exports = (0..count)
                .map(|index| export(&format!("f{index}"), index))
                .collect();
            let instance = InstanceExports::new(exports).unwrap();
            let requests: Vec<_> = (0..count)
                .rev()
                .map(|index| import(&format!("f{index}"), index + count))
                .collect();
            let links = collect_instance_function_links(&[instance], &requests).unwrap();
            for (link, request) in links.iter().zip(&requests) {
                assert_eq!(link.import, request.import);
                assert_eq!(link.supplied, request.import - count);
                assert_eq!(link.supplied_type, link.supplied + 1);
                assert_eq!(link.expected_type, 1);
            }
        }
    }

    #[test]
    fn invalid_instance_export_and_function_type_are_total_errors() {
        let instance = InstanceExports::new(vec![export("value", 0)]).unwrap();
        let mut request = import("value", 1);
        request.instance = 1;
        assert!(collect_instance_function_links(&[instance.clone()], &[request]).is_err());
        assert!(collect_instance_function_links(&[instance], &[import("missing", 1)]).is_err());
        assert!(InstanceExports::new(vec![export("same", 0), export("same", 1)]).is_err());
        let mut untyped = export("untyped", 0);
        untyped.function_type = None;
        assert!(InstanceExports::new(vec![untyped]).is_err());
        for kind in [
            ComponentItemKind::Module,
            ComponentItemKind::Value,
            ComponentItemKind::Type,
            ComponentItemKind::Component,
            ComponentItemKind::Instance,
        ] {
            let instance = InstanceExports::new(vec![InstanceExport {
                name: Arc::from("value"),
                kind,
                index: 0,
                function_type: None,
            }])
            .unwrap();
            assert!(collect_instance_function_links(&[instance], &[import("value", 1)]).is_err());
        }
    }

    #[test]
    fn shared_exports_do_not_share_instance_bindings() {
        let first = InstanceExports::new(vec![export("value", 4)]).unwrap();
        let snapshot = first.clone();
        assert!(Arc::ptr_eq(&first.exports, &snapshot.exports));
        let second = InstanceExports::new(vec![export("value", 9)]).unwrap();
        let mut request = import("value", 0);
        let links =
            collect_instance_function_links(&[first, second.clone()], &[request.clone()]).unwrap();
        assert_eq!(links[0].supplied, 4);
        request.instance = 1;
        let links = collect_instance_function_links(&[snapshot, second], &[request]).unwrap();
        assert_eq!(links[0].supplied, 9);
    }

    #[test]
    fn interleaved_instance_requests_preserve_binding_order_and_identity() {
        let first = InstanceExports::new(
            (0..32)
                .map(|index| export(&format!("f{index}"), index))
                .collect(),
        )
        .unwrap();
        let second = InstanceExports::new(
            (0..32)
                .map(|index| export(&format!("f{index}"), index + 100))
                .collect(),
        )
        .unwrap();
        let requests: Vec<_> = [(0, "f31"), (1, "f7"), (0, "f31"), (1, "f0")]
            .into_iter()
            .enumerate()
            .map(|(slot, (instance, name))| {
                let mut request = import(name, slot as u32 + 200);
                request.instance = instance;
                request
            })
            .collect();
        let links = collect_instance_function_links(&[first, second], &requests).unwrap();
        assert_eq!(
            links.iter().map(|link| link.supplied).collect::<Vec<_>>(),
            vec![31, 107, 31, 100]
        );
        assert_eq!(
            links.iter().map(|link| link.import).collect::<Vec<_>>(),
            vec![200, 201, 202, 203]
        );
    }

    #[test]
    fn absent_instance_metadata_borrows_existing_links_without_copying() {
        let existing = [ComponentFuncLink {
            import: 1,
            supplied: 0,
            expected_type: 2,
            supplied_type: 2,
        }];
        for metadata in [None, Some(InstanceLinkMetadata::default())] {
            let result = bind_instance_function_imports(&existing, metadata.as_ref()).unwrap();
            assert!(matches!(result, Cow::Borrowed(_)));
            assert_eq!(result.as_ptr(), existing.as_ptr());
        }
        let metadata = InstanceLinkMetadata {
            instances: vec![InstanceExports::new(vec![export("value", 4)]).unwrap()],
            imports: vec![import("value", 2)],
        };
        let result = bind_instance_function_imports(&existing, Some(&metadata)).unwrap();
        assert!(matches!(result, Cow::Owned(_)));
        assert_eq!(result.len(), 2);
        assert_eq!(result[0], existing[0]);
        assert_eq!(result[1].import, 2);
        assert_eq!(result[1].supplied, 4);
    }

    #[test]
    #[ignore = "manual instance export lookup benchmark"]
    fn benchmark_instance_export_lookup() {
        const REQUESTS: u32 = 8_192;
        const RUNS: usize = 9;
        for count in [4, 8, 16, 32, 64, 128, 1_024] {
            let instance = InstanceExports::new(
                (0..count)
                    .map(|index| export(&format!("f{index}"), index))
                    .collect(),
            )
            .unwrap();
            let imports: Vec<_> = (0..REQUESTS)
                .map(|slot| import(&format!("f{}", (slot * 17) % count), slot + count))
                .collect();
            let instances = [instance];
            let linear = || -> Vec<ComponentFuncLink> {
                imports
                    .iter()
                    .map(|request| {
                        let instance = &instances[request.instance as usize];
                        let export = instance
                            .exports
                            .iter()
                            .find(|export| export.name == request.name)
                            .unwrap();
                        assert_eq!(export.kind, ComponentItemKind::Func);
                        ComponentFuncLink {
                            import: request.import,
                            supplied: export.index,
                            expected_type: request.expected_type,
                            supplied_type: export.function_type.unwrap(),
                        }
                    })
                    .collect()
            };
            assert_eq!(
                collect_instance_function_links(&instances, &imports).unwrap(),
                linear()
            );
            let mut times = [Vec::new(), Vec::new()];
            for run in 0..RUNS {
                for mode in [run % 2, 1 - run % 2] {
                    let start = std::time::Instant::now();
                    let links = if mode == 0 {
                        linear()
                    } else {
                        collect_instance_function_links(&instances, &imports).unwrap()
                    };
                    std::hint::black_box(links);
                    times[mode].push(start.elapsed());
                }
            }
            for samples in &mut times {
                samples.sort_unstable();
            }
            println!(
                "instance exports={count}, requests={REQUESTS}: scan {:.3} ms, cached {:.3} ms",
                times[0][RUNS / 2].as_secs_f64() * 1000.0,
                times[1][RUNS / 2].as_secs_f64() * 1000.0
            );
        }
    }
}
