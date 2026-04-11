//! Top of the .NET class hierarchy.
//!
//! `Object → MarshalByRefObject → Component`. All three are abstract bases
//! with no widget backing and (at this layer) no setter properties of
//! their own. They exist so the .NET inheritance chain is real:
//!
//! - **Object** — root of every CLR type. Allocates the underlying struct.
//!   Every other .NET class transitively inherits from it.
//!
//! - **MarshalByRefObject** — historical .NET base for cross-AppDomain
//!   marshalling. Real .NET puts it between `Object` and `Component`. We
//!   keep it because the user explicitly asked for the unflattened chain.
//!
//! - **Component** — base for any class that supports the
//!   `IContainer`/`ISite` model (controls, timers, binding sources, …).
//!   In real .NET this is where `Dispose()` lives.
//!
//! No properties, no widget host fn — Phase B exists so Phase C
//! (`Control`) has a real parent to call.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Object",
            parent: None,
            properties: &[],
            widget_host_fn: None,
        },
        DotnetClass {
            name: "MarshalByRefObject",
            parent: Some("Object"),
            properties: &[],
            widget_host_fn: None,
        },
        DotnetClass {
            name: "Component",
            parent: Some("MarshalByRefObject"),
            properties: &[],
            widget_host_fn: None,
        },
    ]
}
