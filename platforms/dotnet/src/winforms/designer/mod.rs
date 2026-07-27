pub mod encoding;
pub mod project;
pub mod resources;
pub mod serialization;
pub mod winforms;

// Data model moved to `crate::winforms`; re-exported here so existing
// `…::projects::vbforms::{Form, Control, …}` paths keep resolving.
pub use crate::winforms::{control, errors, events, form, properties};
// The data model itself (`Form`/`Control`/`ControlType`/`EventType`/
// `PropertyValue`) — one designer path resolves everything a consumer needs.
pub use crate::winforms::control::*;
pub use crate::winforms::errors::*;
pub use crate::winforms::events::*;
pub use crate::winforms::form::*;
pub use crate::winforms::properties::*;

pub use encoding::*;
pub use project::*;
pub use resources::*;
pub use serialization::*;
pub use winforms::*;
