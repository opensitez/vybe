//! `Comparer` / `EqualityComparer` / `StringComparer` as tree types.
//!
//! Declaring these is what lets `Comparer.Default.Compare(a, b)` resolve by an
//! ordinary member walk. Before, `Default` answered a bare string and each .NET
//! frontend recognised that string and rewrote the call itself — so the rule
//! lived in C# and again in VB, and PowerShell had neither.
//!
//! The statics still ANSWER the sentinel string; what is new is that they
//! answer it with a declared TYPE, so the next hop has something to resolve
//! against. That is the same `member_returns` mechanism every other chaining
//! .NET member uses.

use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, MethodBody, MethodDef};

fn comparer_members(class: ClassType) -> ClassType {
    class
        .with_method(MethodDef::new(
            "Compare",
            2,
            MethodBody::Common("dotnet.comparer_compare".into()),
        ))
        .with_method(MethodDef::new(
            "Equals",
            2,
            MethodBody::Common("dotnet.comparer_equals".into()),
        ))
        .with_method(MethodDef::new(
            "GetHashCode",
            1,
            MethodBody::Common("dotnet.comparer_get_hash_code".into()),
        ))
}

fn stat(name: &'static str, common: &'static str) -> MethodDef {
    // Arity 0 — a zero-arg static IS the property read.
    MethodDef::static_method(name, 0, MethodBody::Common(common.into()))
}

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            comparer_members(
                ClassType::new("Comparer")
                    .with_method(stat("Default", "dotnet.comparer_default"))
                    // `Create(comparison)` answers the comparison ITSELF. A
                    // `Comparer<T>` built from a lambda has no state beyond
                    // that lambda, and every consumer here invokes it as
                    // `cmp(a, b)` — wrapping it would add a layer with nothing
                    // in it.
                    .with_method(MethodDef::static_method(
                        "Create",
                        1,
                        MethodBody::Common("collections.identity".into()),
                    )),
            ),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            comparer_members(
                ClassType::new("EqualityComparer")
                    .with_method(stat("Default", "dotnet.equality_comparer_default")),
            ),
        ),
        // `StringComparer` is `System.StringComparer`, not a Collections type.
        //
        // Only the two ORDINAL comparers are declared. .NET also ships
        // `CurrentCulture`/`InvariantCulture` variants, and they are NOT the
        // same rule — culture-aware collation orders "ä" against "a"
        // differently from a code-unit comparison. Declaring them as aliases of
        // the ordinal ones would answer plausibly and wrongly, so the gap stays
        // visible instead.
        DotnetClassExport::new(
            "dotnet.System",
            comparer_members(
                ClassType::new("StringComparer")
                    .with_method(stat("Ordinal", "dotnet.string_comparer_ordinal"))
                    .with_method(stat(
                        "OrdinalIgnoreCase",
                        "dotnet.string_comparer_ordinal_ignore_case",
                    )),
            ),
        ),
    ]
}
