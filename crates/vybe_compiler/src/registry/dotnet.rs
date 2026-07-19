//! RETIRED. The .NET BCL wrapper classes no longer emit a per-class
//! constructor prelude.
//!
//! What replaced it:
//! - **Properties & methods** resolve through the component descriptor at each
//!   call site (`platforms/dotnet` `DotnetSurface::lookup_instance_property` /
//!   `lookup_instance_method`), reached via
//!   `Compiler::dotnet_framework_instance_method_owner`.
//! - **Construction**: control leaves via the GUI-direct `vybe:gui` path
//!   (`emit_new_control` / `try_emit_framework_control_base`); value & drawing
//!   objects (`Point`, `Pen`, `Graphics`, …) via a descriptor constructor
//!   (`ConstructorDef::with_backing`); abstract classes are never directly
//!   constructed.
//! - **Drawing `Body` methods** (`Graphics.DrawLine`, …) lower inline at the
//!   call site through `MethodBody::Common("dotnet.drawing.<Name>")` +
//!   `classes::builder::emit_body_inline`, reusing the same `MethodOp` table
//!   that once built thunk chunks.
//! - **`CreateGraphics`** → `MethodBody::Common("dotnet.control_create_graphics")`.
//!
//! This module is intentionally empty; the file is kept as a tombstone so the
//! history of the retirement is discoverable. No `mod dotnet;` references it.
