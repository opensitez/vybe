use super::helpers::load_vb_profile;

#[test]
fn dotnet_profile_flags_are_enabled_for_vb() {
    let profile = load_vb_profile();

    assert!(profile.namespaces.use_dotnet);
    assert!(profile.namespaces.use_dotnet_resolver);
}
