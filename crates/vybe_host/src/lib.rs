pub mod side_effect;
pub mod modules;
pub mod controls;
pub mod namespaces;
pub mod builtin_types;
pub mod stdlib_aliases;

#[cfg(feature = "gui")]
pub mod gui_state;

pub use side_effect::{SideEffect, PropValue, UIEvent, SideEffectQueue};
pub use modules::{register_all, register_with_capabilities, Capabilities, Capability};
#[cfg(feature = "gui")]
pub use modules::{register_all_with_gui, register_with_capabilities_and_gui};
#[cfg(feature = "gui")]
pub use gui_state::GuiState;
pub use modules::database::{test_connection_and_list_tables, fetch_columns_for_query};
pub use namespaces::setup_namespaces;
pub use controls::get_def;
