pub mod project;
pub mod resources;
pub mod serialization;
pub mod winforms;

// Data model moved to `crate::platforms::dotnet::winforms`; re-exported here so existing
// `…::projects::vbforms::{Form, Control, …}` paths keep resolving.
pub use crate::platforms::dotnet::winforms::*;
pub use crate::platforms::dotnet::winforms::{control, errors, events, form, properties};

pub use project::*;
pub use resources::*;
pub use serialization::*;
pub use winforms::*;
