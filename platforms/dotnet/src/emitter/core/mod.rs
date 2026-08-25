pub mod adodb_adapter;
pub mod array_adapter;
pub mod bindingsource_adapter;
pub mod bitconverter_adapter;
pub mod collections_adapter;
pub mod datagrid_adapter;
pub mod component_classes;
pub mod console_adapter;
pub mod convert_adapter;
pub mod datatable_adapter;
pub mod datetime_adapter;
pub mod encoding_adapter;
pub mod environment_adapter;
pub mod exceptions;
pub mod filesystem_adapter;
pub mod financial_adapter;
pub mod format_picture_adapter;
pub mod gc_adapter;
pub mod guid_adapter;
pub mod host_map;
pub mod http_adapter;
pub mod imports;
pub mod char_adapter;
pub mod float_adapter;
pub mod json_adapter;
pub mod culture_adapter;
pub mod key_value_pair_adapter;
pub mod wait_handle_adapter;
pub mod lazy_adapter;
pub mod linq_adapter;
pub mod lowering;
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
pub mod sqlclient_adapter;
pub mod stopwatch_adapter;
pub mod stream_io_adapter;
pub mod string_adapter;
pub mod string_format_adapter;
pub mod stringbuilder_adapter;
pub mod thread_adapter;
pub mod datetime_format_adapter;
pub mod biginteger_adapter;
pub mod binary_io_adapter;
pub mod memory_stream_adapter;
pub mod datetime_parse_adapter;
pub mod timespan_adapter;
pub mod timezone_adapter;
pub mod types;
pub mod uri_adapter;
pub mod vb_dateandtime_adapter;
pub mod version_adapter;
pub mod visualbasic_adapter;
pub mod xml_linq_adapter;

use vybe_runtime::component_model::ComponentDescriptor;

pub use host_map::static_method_mappings;
pub use imports::default_interface_imports;
pub use namespaces::{is_namespace_root, namespace_roots};
pub use types::{capitalize_data_type, is_known_constant, known_constants, known_type_mappings};

pub fn dotnet_core_component_descriptor() -> ComponentDescriptor {
    super::descriptor::build_dotnet_component_descriptor(
        "dotnet_core",
        super::descriptor::DescriptorPartition::Core,
    )
}
