//! WinForms designer data model (`Form`/`Control`/`ControlType`/`PropertyValue`/
//! events). Framework-specific (.NET WinForms), shared by the VB and C#
//! designer codegen. Self-contained (serde/uuid/std). Sibling GUI frameworks
//! live in their own platform crates (e.g. Lazarus/VCL under `platforms/plib`).

pub mod control;
pub mod errors;
pub mod events;
pub mod form;
pub mod form_modules;
pub mod properties;

pub use control::*;
pub use errors::*;
pub use events::*;
pub use form::*;
pub use properties::*;
