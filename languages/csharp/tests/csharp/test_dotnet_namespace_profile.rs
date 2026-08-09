//! C#/VB reach the .NET surface through the shared namespace resolver.
//!
//! These used to assert `profile.namespaces.use_dotnet` — a language-family
//! gate that has been deleted. A bool being true never proved the surface was
//! reachable; it only proved someone had written `true` in a profile. So they
//! now assert the thing the flag stood in for: the language DECLARES the
//! `dotnet` tree mount, and a framework type actually RESOLVES through it.
//!
//! That is a strictly stronger claim, and it is the one the compiler makes —
//! `is_registered_type(&profile.namespaces.type_scopes, name)` scoped by the
//! profile's own mounts is what replaced the gate at every call site.

/// Register the .NET platform's namespace tree. Idempotent (`Once` inside), so
/// every test in the binary can call it without ordering between them.
fn mount_dotnet() {
    vybe_platform_dotnet::register();
}

fn csharp_profile() -> vybe_runtime::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
        .expect("C# profile parse failed")
}

fn vb_profile() -> vybe_runtime::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
        .expect("VB profile parse failed")
}

#[test]
fn csharp_mounts_the_dotnet_tree_and_resolves_framework_types() {
    mount_dotnet();
    let profile = csharp_profile();

    assert!(
        profile
            .namespaces
            .type_scopes
            .iter()
            .any(|s| s.eq_ignore_ascii_case("dotnet")),
        "C# must declare the `dotnet` tree mount — that declaration is what \
         scopes every .NET answer to the languages that actually have it"
    );
    assert!(
        profile.namespaces.source_imports_are_namespaces && profile.uses_namespace_resolver(),
        "C# must use the shared namespace resolver through profile data"
    );

    // The mount is only meaningful if types resolve through it. `Button` is a
    // GUI control and `DateTime` a core BCL type, so this covers both halves
    // of the descriptor (winforms + core).
    for name in ["Button", "Form", "DateTime"] {
        assert!(
            vybe_runtime::namespaces::is_registered_type(&profile.namespaces.type_scopes, name),
            "`{name}` must resolve as a registered type through C#'s declared \
             scopes — this is the question that replaced the `use_dotnet` gate"
        );
    }
}

#[test]
fn vb_and_csharp_both_use_shared_dotnet_namespace_system() {
    mount_dotnet();
    let vb = vb_profile();
    let cs = csharp_profile();

    for (lang, profile) in [("VB", &vb), ("C#", &cs)] {
        assert!(
            profile
                .namespaces
                .type_scopes
                .iter()
                .any(|s| s.eq_ignore_ascii_case("dotnet")),
            "{lang} must mount the shared `dotnet` tree, not its own registration"
        );
        assert!(
            vybe_runtime::namespaces::is_registered_type(&profile.namespaces.type_scopes, "Button"),
            "{lang} must resolve `Button` through that one shared mount"
        );
    }

    assert!(
        vb.namespaces.source_imports_are_namespaces
            && cs.namespaces.source_imports_are_namespaces
            && vb.uses_namespace_resolver()
            && cs.uses_namespace_resolver(),
        "VB and C# must both use shared namespace resolver semantics"
    );
}

#[test]
fn a_language_without_the_mount_does_not_resolve_framework_types() {
    mount_dotnet();

    // The half the old assertions could not express, and the reason the gate
    // existed at all: `canonical_control_name` is a shared table holding
    // generic words (`image`, `panel`, `label`, `timer`), so scoping is the
    // ONLY thing stopping a Python `class X(Timer)` becoming a GUI control.
    // With the tree registered process-wide, an empty scope must still answer
    // no — otherwise every gate removal that relies on this is unsound.
    let no_scopes: Vec<String> = Vec::new();
    for name in ["Button", "Timer", "Panel", "Label"] {
        assert!(
            !vybe_runtime::namespaces::is_registered_type(&no_scopes, name),
            "`{name}` must NOT resolve for a language that mounts no tree, even \
             though the .NET platform is linked into this binary"
        );
    }
}
