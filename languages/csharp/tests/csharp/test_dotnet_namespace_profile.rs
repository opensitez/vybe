#[test]
fn csharp_profile_enables_dotnet_namespace_resolution() {
    let profile = vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
        .expect("C# profile parse failed");

    assert!(
        profile.namespaces.use_dotnet,
        "C# must keep .NET namespace support enabled"
    );
    assert!(
        profile.namespaces.use_dotnet_resolver,
        "C# must keep shared .NET dotted-name resolver enabled"
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
        vb.namespaces.use_dotnet_resolver && cs.namespaces.use_dotnet_resolver,
        "VB and C# must both use shared .NET resolver semantics"
    );
}
