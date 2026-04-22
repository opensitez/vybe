pub mod imports;
pub mod component_classes;
pub mod host_map;
pub mod types;
pub mod namespaces;

use vybe_bytecode::component_model::ComponentDescriptor;

pub use imports::default_interface_imports;
pub use host_map::{map_host_func, namespace_to_host_module, static_method_mappings};
pub use types::{capitalize_data_type, is_known_constant, known_constants, known_type_mappings};
pub use namespaces::{is_namespace_root, namespace_roots};

pub fn dotnet_core_component_descriptor() -> ComponentDescriptor {
    super::descriptor::build_dotnet_component_descriptor(
        "dotnet_core",
        super::descriptor::DescriptorPartition::Core,
    )
}