pub mod c;
pub mod cobol;
pub mod csharp;
pub mod dart;
pub mod form_modules;
pub mod fortran;
pub mod go;
pub mod js;
pub mod lua;
pub mod pascal;
pub mod php;
pub mod python;
pub mod ruby;
pub mod vb;
pub mod wast;

use crate::ast::Module;
use vybe_bytecode::Chunk;

/// Per-language emit dispatcher: routes a `common:<prefix>.<op>` name to
/// the language's own emitter, where `<prefix>` is the language name.
/// Runs the matching arm as a side-effect and returns `true`; returns
/// `false` for names it doesn't own (so the caller can fall through).
/// Registered per-language via [`Language::emit_dispatch`] — the same
/// way file extensions are registered — so adding a language never
/// touches the central `emitter::dispatch`.
pub type EmitDispatch = fn(&str, &mut Vec<Chunk>, usize, u8, u32) -> bool;

/// A registered language in vybex.
pub struct Language {
    /// Canonical language name — matches the profile's `[info].name` and
    /// the per-language sub-module name. Use [`find_by_name`] to look
    /// up; this field exists so dispatch by name doesn't require parsing
    /// the TOML profile each time.
    pub name: &'static str,
    /// Parse function: source → common AST
    pub parse: fn(&str) -> Result<Module, String>,
    /// Embedded profile TOML source
    pub profile_source: fn() -> &'static str,
    /// Emit dispatcher for this language's `common:<name>.*` ops, or
    /// `None` if the language only uses shared common ops (collections,
    /// strings, ...) and a platform (e.g. `dotnet`). See [`EmitDispatch`].
    pub emit_dispatch: Option<EmitDispatch>,
}

/// All registered languages. This is the ONLY place you add a new language.
/// The profile's [info] section defines the name, extensions, and other metadata.
pub fn all() -> Vec<Language> {
    vec![
        Language {
            name: "c",
            parse: c::parse,
            profile_source: c::profile_source,
            emit_dispatch: None,
        },
        Language {
            name: "vb",
            parse: vb::parse,
            profile_source: vb::profile_source,
            emit_dispatch: Some(vb::emitter::dispatch::dispatch),
        },
        Language {
            name: "js",
            parse: js::parse,
            profile_source: js::profile_source,
            emit_dispatch: Some(js::emitter::dispatch::dispatch),
        },
        Language {
            name: "lua",
            parse: lua::parse,
            profile_source: lua::profile_source,
            emit_dispatch: Some(lua::emitter::dispatch::dispatch),
        },
        Language {
            name: "pascal",
            parse: pascal::parse,
            profile_source: pascal::profile_source,
            emit_dispatch: Some(pascal::emitter::dispatch::dispatch),
        },
        Language {
            name: "csharp",
            parse: csharp::parse,
            profile_source: csharp::profile_source,
            emit_dispatch: None,
        },
        Language {
            name: "python",
            parse: python::parse,
            profile_source: python::profile_source,
            emit_dispatch: Some(python::emitter::dispatch::dispatch),
        },
        Language {
            name: "php",
            parse: php::parse,
            profile_source: php::profile_source,
            emit_dispatch: Some(php::emitter::dispatch::dispatch),
        },
        Language {
            name: "ruby",
            parse: ruby::parse,
            profile_source: ruby::profile_source,
            emit_dispatch: Some(ruby::emitter::dispatch::dispatch),
        },
        Language {
            name: "dart",
            parse: dart::parse,
            profile_source: dart::profile_source,
            emit_dispatch: Some(dart::emitter::dispatch::dispatch),
        },
        Language {
            name: "cobol",
            parse: cobol::parse,
            profile_source: cobol::profile_source,
            emit_dispatch: Some(cobol::emitter::dispatch::dispatch),
        },
        Language {
            name: "fortran",
            parse: fortran::parse,
            profile_source: fortran::profile_source,
            emit_dispatch: Some(fortran::emitter::dispatch::dispatch),
        },
        Language {
            name: "go",
            parse: go::parse,
            profile_source: go::profile_source,
            emit_dispatch: Some(go::emitter::dispatch::dispatch),
        },
        Language {
            name: "wast",
            parse: wast::parse,
            profile_source: wast::profile_source,
            emit_dispatch: None,
        },
    ]
}

/// Resolve the emit dispatcher that owns `prefix` in a `common:<prefix>.*`
/// name. Languages register theirs via [`Language::emit_dispatch`]; shared
/// platforms (currently only `dotnet`, used by VB/C#/JS) resolve through
/// [`crate::emitter::platform_emit_dispatch`]. Returns `None` for the
/// genuinely-shared prefixes (collections, dict, strings, ...) which the
/// central dispatcher owns directly.
pub fn emit_dispatch_for(prefix: &str) -> Option<EmitDispatch> {
    if let Some(d) = all()
        .into_iter()
        .find(|l| l.name == prefix)
        .and_then(|l| l.emit_dispatch)
    {
        return Some(d);
    }
    crate::emitter::platform_emit_dispatch(prefix)
}

/// Find a language by canonical name (e.g. `"js"`, `"php"`, `"python"`).
/// Returns the registered [`Language`] entry or None if no match.
pub fn find_by_name(name: &str) -> Option<Language> {
    all().into_iter().find(|l| l.name == name)
}

/// Find a language by file extension. Reads [info].extensions from each profile.
pub fn find_by_extension(ext: &str) -> Option<Language> {
    let ext_lower = ext.to_lowercase();
    for lang in all() {
        if let Some(extensions) = read_extensions((lang.profile_source)()) {
            if extensions.iter().any(|e| e == &ext_lower) {
                return Some(lang);
            }
        }
    }
    None
}

/// List all supported extensions (for usage message).
pub fn supported_extensions() -> Vec<String> {
    let mut exts = Vec::new();
    for lang in all() {
        if let Some(extensions) = read_extensions((lang.profile_source)()) {
            for e in extensions {
                exts.push(e);
            }
        }
    }
    exts
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
