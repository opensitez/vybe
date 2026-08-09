use super::helpers::load_vb_profile;

/// VB reaches the .NET surface through the shared namespace resolver.
///
/// This used to assert `profile.namespaces.use_dotnet`, a language-family gate
/// that has been deleted. A bool being true never proved the surface was
/// reachable — it proved someone wrote `true` in a profile. The assertion is
/// now the thing the flag stood in for: VB DECLARES the `dotnet` tree mount,
/// and a framework type actually RESOLVES through it. That is the same
/// question the compiler asks (`is_registered_type` scoped by `type_scopes`).
#[test]
fn dotnet_surface_is_reachable_for_vb() {
    vybe_platform_dotnet::register();
    let profile = load_vb_profile();

    assert!(
        profile
            .namespaces
            .type_scopes
            .iter()
            .any(|s| s.eq_ignore_ascii_case("dotnet")),
        "VB must declare the `dotnet` tree mount"
    );
    assert!(profile.namespaces.source_imports_are_namespaces);
    assert!(profile.uses_namespace_resolver());

    for name in ["Button", "Form", "DateTime"] {
        assert!(
            vybe_runtime::namespaces::is_registered_type(&profile.namespaces.type_scopes, name),
            "`{name}` must resolve as a registered type through VB's declared scopes"
        );
    }
}
