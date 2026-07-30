#[test]
fn csharp_profile_enables_dotnet_namespace_resolution() {
    let profile = vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
        .expect("C# profile parse failed");

    assert!(
        profile.namespaces.use_dotnet,
        "C# must keep .NET namespace support enabled"
    );
    assert!(
        profile.namespaces.source_imports_are_namespaces && profile.uses_namespace_resolver(),
        "C# must use the shared namespace resolver through profile data"
    );
}

#[test]
fn vb_and_csharp_both_use_shared_dotnet_namespace_system() {
    let vb = vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
        .expect("VB profile parse failed");
    let cs = vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
        .expect("C# profile parse failed");

    assert!(
        vb.namespaces.use_dotnet && cs.namespaces.use_dotnet,
        "VB and C# must both use shared .NET namespace registration"
    );
    assert!(
        vb.namespaces.source_imports_are_namespaces
            && cs.namespaces.source_imports_are_namespaces
            && vb.uses_namespace_resolver()
            && cs.uses_namespace_resolver(),
        "VB and C# must both use shared namespace resolver semantics"
    );
}
