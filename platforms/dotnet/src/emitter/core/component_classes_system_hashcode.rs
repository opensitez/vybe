//! `System.HashCode` — the type declaration; `hashcode_adapter` is the content.
//!
//! Its own file rather than a block inside `component_classes_system.rs` for
//! the reason `component_classes_system_version.rs` is: a type whose whole
//! surface is one adapter reads better beside that adapter than inside a
//! 2000-line list.
//!
//! ⚠ `Combine` is declared ONCE though .NET declares it at eight arities. The
//! static lookup in this platform ignores arity (a known defect, recorded in
//! the plan), so eight rows would resolve identically; the adapter reads `argc`
//! instead. If that lookup ever grows arity awareness, this needs the other
//! seven rows — the adapter already handles them.

use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![DotnetClassExport::new(
        "dotnet.System",
        ClassType::new("HashCode")
            .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.hashcode_new"))
            .with_method(MethodDef::static_method(
                "Combine",
                2,
                MethodBody::Common("dotnet.hashcode_combine".into()),
            ))
            .with_method(MethodDef::new(
                "Add",
                1,
                MethodBody::Common("dotnet.hashcode_add".into()),
            ))
            .with_method(MethodDef::new(
                "Add",
                2,
                MethodBody::Common("dotnet.hashcode_add_comparer".into()),
            ))
            .with_method(MethodDef::new(
                "ToHashCode",
                0,
                MethodBody::Common("dotnet.hashcode_to_hash_code".into()),
            ))
            // Both throw `NotSupportedException` in .NET — declared so they
            // throw here rather than resolving to `Object`'s versions and
            // quietly answering a number.
            .with_method(MethodDef::new(
                "GetHashCode",
                0,
                MethodBody::Common("dotnet.hashcode_unsupported".into()),
            ))
            .with_method(MethodDef::new(
                "Equals",
                1,
                MethodBody::Common("dotnet.hashcode_unsupported".into()),
            )),
    )]
}
