use std::sync::LazyLock;

use super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::ClassType;

#[path = "component_classes_collections.rs"]
mod component_classes_collections;
#[path = "component_classes_common.rs"]
mod component_classes_common;
#[path = "component_classes_data_drawing.rs"]
mod component_classes_data_drawing;
#[path = "component_classes_diagnostics_process.rs"]
mod component_classes_diagnostics_process;
#[path = "component_classes_io.rs"]
mod component_classes_io;
#[path = "component_classes_linq.rs"]
mod component_classes_linq;
#[path = "component_classes_network.rs"]
mod component_classes_network;
#[path = "component_classes_span.rs"]
mod component_classes_span;
#[path = "component_classes_system.rs"]
mod component_classes_system;
#[path = "component_classes_system_values.rs"]
mod component_classes_system_values;
#[path = "component_classes_system_version.rs"]
mod component_classes_system_version;
#[path = "component_classes_text.rs"]
mod component_classes_text;
#[path = "component_classes_threading.rs"]
mod component_classes_threading;
#[path = "component_classes_uri.rs"]
mod component_classes_uri;
#[path = "component_classes_xml.rs"]
mod component_classes_xml;

pub fn class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        let mut exports = component_classes_collections::exports();
        component_classes_linq::apply_linq_registrations(&mut exports);
        // The `System.Linq.Enumerable` surface, declared once. Any enumerable
        // receiver falls back to it via `lookup_instance_method` — which also
        // makes it the only home for array extension methods like `AsSpan`.
        let mut enumerable = component_classes_linq::enumerable_export();
        component_classes_span::add_array_extension_methods(&mut enumerable.class);
        exports.push(enumerable);
        exports.push(component_classes_linq::enumerable_static_export());
        exports.extend(component_classes_span::exports());
        exports.extend(component_classes_system::exports());
        exports.extend(component_classes_system_values::exports());
        exports.extend(component_classes_system_version::exports());
        exports.extend(component_classes_threading::exports());
        exports.extend(component_classes_text::exports());
        exports.extend(component_classes_data_drawing::exports());
        exports.extend(component_classes_diagnostics_process::exports());
        exports.extend(component_classes_network::exports());
        exports.extend(component_classes_uri::exports());
        exports.extend(component_classes_io::exports());
        exports.extend(component_classes_xml::exports());
        exports
    });
    EXPORTS.as_slice()
}

pub fn component_class_exports() -> &'static [(&'static str, ClassType)] {
    static EXPORTS: LazyLock<Vec<(&'static str, ClassType)>> = LazyLock::new(|| {
        class_exports()
            .iter()
            .map(|export| (export.interface, export.class.clone()))
            .collect()
    });
    EXPORTS.as_slice()
}
