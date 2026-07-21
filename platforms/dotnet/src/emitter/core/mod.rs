pub mod adodb_adapter;
pub mod array_adapter;
pub mod collections_adapter;
pub mod component_classes;
pub mod console_adapter;
pub mod datatable_adapter;
pub mod datetime_adapter;
pub mod environment_adapter;
pub mod filesystem_adapter;
pub mod format_picture_adapter;
pub mod guid_adapter;
pub mod host_map;
pub mod imports;
pub mod linq_adapter;
pub mod namespaces;
pub mod node_socket_adapter;
pub mod numeric_format;
pub mod oledb_adapter;
pub mod parse_adapter;
pub mod process_adapter;
pub mod random_adapter;
pub mod regex_adapter;
pub mod runtime_adapter;
pub mod sockets_adapter;
pub mod span_adapter;
pub mod stopwatch_adapter;
pub mod stream_io_adapter;
pub mod string_format_adapter;
pub mod stringbuilder_adapter;
pub mod thread_adapter;
pub mod timespan_adapter;
pub mod types;
pub mod version_adapter;
pub mod xml_linq_adapter;

use vybe_bytecode::component_model::ComponentDescriptor;

pub use host_map::{map_host_func, namespace_to_host_module, static_method_mappings};
pub use imports::default_interface_imports;
pub use namespaces::{is_namespace_root, namespace_roots};
pub use types::{capitalize_data_type, is_known_constant, known_constants, known_type_mappings};

pub fn dotnet_core_component_descriptor() -> ComponentDescriptor {
    super::descriptor::build_dotnet_component_descriptor(
        "dotnet_core",
        super::descriptor::DescriptorPartition::Core,
    )
}
