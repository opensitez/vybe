pub mod imports;
pub mod component_classes;
pub mod host_map;
pub mod namespaces;
pub mod types;

use vybe_bytecode::component_model::ComponentDescriptor;

#[path = "../classes/mod.rs"]
pub mod classes;

pub use imports::default_interface_imports;
pub use host_map::{map_host_func, namespace_to_host_module, static_method_mappings};
pub use namespaces::{is_namespace_root, namespace_roots};
pub use types::{capitalize_control_name, is_noop_method, known_type_mappings, noop_methods};

pub fn dotnet_winforms_component_descriptor() -> ComponentDescriptor {
    super::descriptor::build_dotnet_component_descriptor(
        "dotnet_winforms",
        super::descriptor::DescriptorPartition::WinForms,
    )
}