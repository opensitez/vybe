pub mod cobol;
pub mod csharp;
pub mod dart;
pub mod form_modules;
pub mod fortran;
pub mod js;
pub mod pascal;
pub mod php;
pub mod python;
pub mod ruby;
pub mod vb;

use crate::ast::Module;

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
}

/// All registered languages. This is the ONLY place you add a new language.
/// The profile's [info] section defines the name, extensions, and other metadata.
pub fn all() -> Vec<Language> {
    vec![
        Language { name: "vb",      parse: vb::parse,      profile_source: vb::profile_source },
        Language { name: "js",      parse: js::parse,      profile_source: js::profile_source },
        Language { name: "pascal",  parse: pascal::parse,  profile_source: pascal::profile_source },
        Language { name: "csharp",  parse: csharp::parse,  profile_source: csharp::profile_source },
        Language { name: "python",  parse: python::parse,  profile_source: python::profile_source },
        Language { name: "php",     parse: php::parse,     profile_source: php::profile_source },
        Language { name: "ruby",    parse: ruby::parse,    profile_source: ruby::profile_source },
        Language { name: "dart",    parse: dart::parse,    profile_source: dart::profile_source },
        Language { name: "cobol",   parse: cobol::parse,   profile_source: cobol::profile_source },
        Language { name: "fortran", parse: fortran::parse, profile_source: fortran::profile_source },
    ]
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
    Some(arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
}
