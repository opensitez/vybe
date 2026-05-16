use std::sync::LazyLock;

use vybe_bytecode::component_model::ClassType;

use super::winforms::classes::DotnetClass;

#[derive(Debug, Clone)]
pub struct DotnetClassExport {
    pub interface: &'static str,
    pub class: ClassType,
    pub wrapper: Option<DotnetClass>,
}

impl DotnetClassExport {
    pub fn new(interface: &'static str, class: ClassType) -> Self {
        Self {
            interface,
            class,
            wrapper: None,
        }
    }

    pub fn with_wrapper(interface: &'static str, class: ClassType, wrapper: DotnetClass) -> Self {
        Self {
            interface,
            class,
            wrapper: Some(wrapper),
        }
    }
}

pub fn dotnet_class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        let mut exports = Vec::new();
        exports.extend_from_slice(super::core::component_classes::class_exports());
        exports.extend_from_slice(super::winforms::component_classes::class_exports());
        exports
    });
    EXPORTS.as_slice()
}