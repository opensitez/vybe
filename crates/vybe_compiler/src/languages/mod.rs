// `java` now lives in the `languages/java` crate (`vybe_lang_java`); it
// registers itself through `vybe_runtime::registry` via the aggregator.

use vybe_runtime::Chunk;

/// Per-language emit dispatcher: routes a `common:<prefix>.<op>` name to
/// the language's own emitter, where `<prefix>` is the language name.
/// Runs the matching arm as a side-effect and returns `true`; returns
/// `false` for names it doesn't own (so the caller can fall through).
/// Registered per-language via [`Language::emit_dispatch`] — the same
/// way file extensions are registered — so adding a language never
/// touches the central `emitter::dispatch`.
pub type EmitDispatch = fn(&str, &mut Vec<Chunk>, usize, u8, u32) -> bool;

/// A registered language — the shared plugin descriptor from `vybe_plugin`.
/// Built-in languages register through [`crate::ensure_languages_registered`];
/// extracted language crates (e.g. `vybe_lang_java`) register themselves.
pub type Language = vybe_runtime::registry::LanguageDef;

/// All registered languages. This is the ONLY place you add a new language.
/// The profile's [info] section defines the name, extensions, and other metadata.
pub fn all() -> Vec<Language> {
    crate::ensure_languages_registered();
    vybe_runtime::registry::all()
}

/// Resolve the emit dispatcher that owns `prefix` in a `common:<prefix>.*`
/// name. Languages register theirs via [`Language::emit_dispatch`]; shared
/// platforms (currently only `dotnet`, used by VB/C#/JS) resolve through
/// [`crate::primitives::platform_emit_dispatch`]. Returns `None` for the
/// genuinely-shared prefixes (collections, dict, strings, ...) which the
/// central dispatcher owns directly.
pub fn emit_dispatch_for(prefix: &str) -> Option<EmitDispatch> {
    crate::ensure_languages_registered();
    vybe_runtime::registry::emit_dispatch_for(prefix)
        .or_else(|| crate::primitives::platform_emit_dispatch(prefix))
}

/// Find a language by canonical name (e.g. `"js"`, `"php"`, `"python"`).
/// Returns the registered [`Language`] entry or None if no match.
pub fn find_by_name(name: &str) -> Option<Language> {
    all().into_iter().find(|l| l.name == name)
}

/// `extension → language name`, derived from every registered profile's
/// `[info].extensions` and then REUSED.
///
/// What a `.kt` file resolves to does not change between programs, but this
/// used to be re-derived on every lookup: walk all registered languages,
/// `toml::from_str` each COMPLETE profile (the largest is 1,243 lines), read
/// one array from it. Under the warm worker that landed once per job and cost
/// **36% of the time to run a one-line program** — measured with `sample`, the
/// single largest item in the profile.
///
/// Ordered, so the "Supported: …" error message and first-declarer-wins are
/// both stable rather than hash-order.
///
/// Rebuilt when the registered-language COUNT changes rather than built once by
/// a bare `OnceLock`: languages register through the plugin registry, and one
/// registering after the first lookup would otherwise stay invisible for the
/// life of the process.
fn with_extension_map<R>(f: impl FnOnce(&std::collections::BTreeMap<String, String>) -> R) -> R {
    use std::collections::BTreeMap;
    use std::sync::{Mutex, OnceLock};
    static MAP: OnceLock<Mutex<(usize, BTreeMap<String, String>)>> = OnceLock::new();
    let cell = MAP.get_or_init(|| Mutex::new((usize::MAX, BTreeMap::new())));
    let langs = all();
    let mut guard = cell.lock().unwrap();
    if guard.0 != langs.len() {
        let mut map = BTreeMap::new();
        for lang in &langs {
            for ext in read_extensions((lang.profile_source)()).unwrap_or_default() {
                // First declarer wins — the same precedence the linear scan
                // had, which stopped at the first language claiming the
                // extension.
                map.entry(ext).or_insert_with(|| lang.name.to_string());
            }
        }
        *guard = (langs.len(), map);
    }
    f(&guard.1)
}

/// Find a language by file extension. Reads [info].extensions from each profile.
pub fn find_by_extension(ext: &str) -> Option<Language> {
    let name = with_extension_map(|m| m.get(&ext.to_lowercase()).cloned())?;
    find_by_name(&name)
}

/// List all supported extensions (for usage message).
pub fn supported_extensions() -> Vec<String> {
    with_extension_map(|m| m.keys().cloned().collect())
}

/// Read the extensions array from a profile's [info] section.
fn read_extensions(profile_src: &str) -> Option<Vec<String>> {
    let root: toml::Value = toml::from_str(profile_src).ok()?;
    let info = root.get("info")?;
    let arr = info.get("extensions")?.as_array()?;
    Some(
        arr.iter()
            .filter_map(|v| v.as_str().map(|s| s.to_string()))
            .collect(),
    )
}
