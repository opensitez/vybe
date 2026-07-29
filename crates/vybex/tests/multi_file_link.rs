//! Command-line multi-file linking — `vybex main.c util.c`, C-compiler style.
//!
//! `projects::load_many` is the entry point the CLI uses for its positional
//! arguments. One path keeps the previous single-file behaviour exactly;
//! several paths link into ONE multi-source `Bundle` (the same shape the
//! `.vybe`/`.vbproj` project loaders produce), with the first path as the
//! entry file.

use std::path::PathBuf;
use std::sync::Once;
use vybe_compiler::projects;

/// Languages resolve by extension through the process-global plugin registry,
/// which is populated by the single plugin loop. Run it once for the suite.
fn languages_registered() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut vm = vybe_runtime::VM::new();
        vybex::cli::register_plugins(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    });
}

/// Write `files` into a fresh temp dir and return their paths, in order.
fn scratch(name: &str, files: &[(&str, &str)]) -> (PathBuf, Vec<PathBuf>) {
    let dir = std::env::temp_dir().join(format!("vybe_multi_file_{name}"));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("create scratch dir");
    let paths = files
        .iter()
        .map(|(file, code)| {
            let path = dir.join(file);
            std::fs::write(&path, code).expect("write source");
            path
        })
        .collect();
    (dir, paths)
}

#[test]
fn several_source_files_link_into_one_bundle() {
    languages_registered();
    let (_dir, paths) = scratch(
        "link",
        &[
            ("main.js", "console.log(triple(14));\n"),
            ("helper.js", "function triple(x) { return x * 3; }\n"),
        ],
    );

    let bundle = projects::load_many(&paths).expect("multiple files must link");
    assert_eq!(bundle.sources.len(), 2, "both files belong to the bundle");
    assert_eq!(bundle.language.name, "js");
    // The entry file is first, and names the bundle.
    assert_eq!(bundle.name, "main");
    assert!(bundle.sources[0].path.ends_with("main.js"));
    assert!(bundle.sources[1].path.ends_with("helper.js"));
}

#[test]
fn single_file_still_loads_as_one_source() {
    languages_registered();
    let (_dir, paths) = scratch("single", &[("solo.js", "console.log(42);\n")]);

    let bundle = projects::load_many(&paths).expect("one file must still load");
    assert_eq!(bundle.sources.len(), 1);
    assert_eq!(bundle.name, "solo");
}

#[test]
fn mixing_languages_in_one_link_is_rejected() {
    // Each front-end lowers through its own walker/profile, so a single link
    // step has no defined semantics across languages — fail loudly instead of
    // silently compiling only the first file.
    languages_registered();
    let (_dir, paths) = scratch(
        "mixed",
        &[
            ("main.js", "console.log(1);\n"),
            ("helper.py", "def helper():\n    return 1\n"),
        ],
    );

    let Err(err) = projects::load_many(&paths) else {
        panic!("mixed languages must fail");
    };
    assert!(
        err.contains("same language"),
        "error should explain the language mismatch, got: {err}"
    );
}

#[test]
fn project_file_cannot_be_combined_with_loose_sources() {
    // A project file already declares its own source list.
    languages_registered();
    let (_dir, paths) = scratch(
        "proj",
        &[
            ("app.vybe", "[project]\nname=\"app\"\n"),
            ("extra.js", "console.log(1);\n"),
        ],
    );

    let Err(err) = projects::load_many(&paths) else {
        panic!("project + loose source must fail");
    };
    assert!(
        err.contains("project file"),
        "error should name the project-file conflict, got: {err}"
    );
}

#[test]
fn unknown_extension_in_a_later_file_is_reported() {
    languages_registered();
    let (_dir, paths) = scratch(
        "unknown",
        &[("main.js", "console.log(1);\n"), ("notes.xyz", "hello\n")],
    );

    let Err(err) = projects::load_many(&paths) else {
        panic!("unknown extension must fail");
    };
    assert!(
        err.contains("xyz"),
        "error should name the offending extension, got: {err}"
    );
}
