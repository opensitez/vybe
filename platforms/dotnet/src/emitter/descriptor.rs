use std::collections::{BTreeMap, HashSet};
use vybe_runtime::component::FuncSig;
use vybe_runtime::component_model::{
    ClassType, ComponentDescriptor, ConstructorTarget, HostTarget, MethodBody,
};

use super::class_exports;

pub(crate) fn build_dotnet_component_descriptor(
    name: &str,
    partition: DescriptorPartition,
) -> ComponentDescriptor {
    let mut descriptor = ComponentDescriptor::new(name);
    let mut seen_host_imports: HashSet<(String, String)> = HashSet::new();
    let mut class_exports: BTreeMap<String, (&'static str, ClassType)> = BTreeMap::new();

    for export in class_exports::dotnet_class_exports() {
        if !partition.accepts_interface(export.interface) {
            continue;
        }
        register_component_class_imports(&mut descriptor, &mut seen_host_imports, &export.class);
        let key = class_export_key(export.interface, &export.class.name);
        class_exports.insert(key, (export.interface, export.class.clone()));
    }

    for (_, (iface, class_type)) in class_exports {
        let export_name = class_type.name.clone();
        descriptor.add_class(class_type.clone());
        descriptor.add_export_class(iface, &export_name, class_type);
    }

    descriptor
}

pub(crate) fn merge_component_descriptor(
    into: &mut ComponentDescriptor,
    other: ComponentDescriptor,
) {
    into.imports.extend(other.imports);
    into.exports.extend(other.exports);
    into.classes.extend(other.classes);
    into.resources.extend(other.resources);
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DescriptorPartition {
    Core,
    WinForms,
}

impl DescriptorPartition {
    fn accepts_interface(self, interface: &str) -> bool {
        match self {
            DescriptorPartition::Core => !is_winforms_interface(interface),
            DescriptorPartition::WinForms => is_winforms_interface(interface),
        }
    }
}

fn is_winforms_interface(interface: &str) -> bool {
    interface == "dotnet.System.Windows.Forms"
}

pub(crate) fn class_export_key(iface: &str, name: &str) -> String {
    format!("{}::{}", iface, name.to_lowercase())
}

fn register_component_class_imports(
    descriptor: &mut ComponentDescriptor,
    seen_host_imports: &mut HashSet<(String, String)>,
    class: &ClassType,
) {
    for property in &class.properties {
        if let Some(setter) = &property.setter {
            register_host_import(descriptor, seen_host_imports, setter, 3);
        }
        if let Some(getter) = &property.getter {
            register_host_import(descriptor, seen_host_imports, getter, 1);
        }
    }

    for method in &class.methods {
        if let MethodBody::HostCall(target) = &method.body {
            register_host_import(descriptor, seen_host_imports, target, method.arity);
        }
    }

    if let Some(constructor) = class.constructor() {
        if let Some(ConstructorTarget::Host(target)) = &constructor.backing {
            register_host_import(descriptor, seen_host_imports, target, constructor.arity);
        }
    }
}

fn register_host_import(
    descriptor: &mut ComponentDescriptor,
    seen_host_imports: &mut HashSet<(String, String)>,
    target: &HostTarget,
    arity: u8,
) {
    let key = (target.module.clone(), target.name.clone());
    if seen_host_imports.insert(key) {
        descriptor.add_import_fn(
            &target.module,
            &target.name,
            FuncSig {
                name: target.name.clone(),
                params: vec![vybe_runtime::component::ValType::Any; arity as usize],
                results: vec![vybe_runtime::component::ValType::Any],
            },
        );
    }
}
