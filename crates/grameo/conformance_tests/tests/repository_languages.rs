//! Opt-in black-box source conformance, batched to avoid deriving eighteen
//! reference parsers in the ordinary development loop. No frontend/compiler
//! dependencies or reference implementation source are used.
#![cfg(any(
    feature = "corpus-scripting",
    feature = "corpus-c-family",
    feature = "corpus-jvm-go-dart",
    feature = "corpus-dynamic",
    feature = "corpus-business",
    feature = "corpus-vb-shell-wast"
))]
use pest::Parser as _;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
};
use grameo::{
    MatchOptions, ParseErrorKind,
    engine::{parse_program, recognize_program},
    program::Program,
};
type Shape = Vec<(String, usize, usize, usize)>;
#[cfg(feature = "corpus-scripting")]
mod python_preprocessor {
    include!(concat!(env!("OUT_DIR"), "/python_preprocessor.rs"));
}
#[cfg(feature = "corpus-business")]
// Preserve the existing frontend helper verbatim for preprocessing parity.
#[allow(clippy::collapsible_if)]
mod fortran_preprocessor {
    include!(concat!(env!("OUT_DIR"), "/fortran_preprocessor.rs"));
}

fn reference_shape<R: pest::RuleType>(pairs: pest::iterators::Pairs<'_, R>) -> Shape {
    fn visit<R: pest::RuleType>(pair: pest::iterators::Pair<'_, R>, depth: usize, out: &mut Shape) {
        out.push((
            format!("{:?}", pair.as_rule()),
            pair.as_span().start(),
            pair.as_span().end(),
            depth,
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut shape = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut shape);
    }
    shape
}
fn native_shape(tree: &grameo::ParseTree<'_, '_>) -> Shape {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((pair.rule_name().to_owned(), span.start, span.end, depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut shape = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut shape);
    }
    shape
}
fn paths(root: &Path, extensions: &[&str], out: &mut Vec<PathBuf>) {
    for entry in fs::read_dir(root).unwrap() {
        let path = entry.unwrap().path();
        if path.is_dir() {
            paths(&path, extensions, out);
        } else if path
            .extension()
            .and_then(|s| s.to_str())
            .is_some_and(|s| extensions.contains(&s))
        {
            out.push(path);
        }
    }
}
fn corpus(language: &str, extensions: &[&str]) -> Vec<(String, String)> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../../tests")
        .join(language);
    let mut files = Vec::new();
    paths(&root, extensions, &mut files);
    files.sort();
    assert!(
        !files.is_empty(),
        "missing source fixtures: {}",
        root.display()
    );
    let mut samples = Vec::new();
    // Spread across the sorted corpus, rather than only one alphabetical category.
    // Filter before sampling so enormous fixtures cannot silently reduce coverage.
    files.retain(|path| fs::metadata(path).unwrap().len() <= 8_192);
    let count = 12.min(files.len());
    assert!(count > 0);
    for index in 0..count {
        let path = &files[index * files.len() / count];
        let source = fs::read_to_string(path).unwrap();
        samples.push((
            path.strip_prefix(&root).unwrap().display().to_string(),
            source,
        ));
    }
    let transformed = samples
        .iter()
        .filter_map(|(label, source)| {
            let processed = match language {
                #[cfg(feature = "corpus-scripting")]
                "python" => Some(python_preprocessor::preprocess_indentation(source)),
                #[cfg(feature = "corpus-business")]
                "fortran" => Some(fortran_preprocessor::rewrite_hollerith(source)),
                _ => {
                    let _ = source;
                    None
                }
            }?;
            Some((format!("preprocessed:{label}"), processed))
        })
        .collect::<Vec<_>>();
    samples.extend(transformed);
    samples
}
fn check<P: Program>(
    program: &P,
    entry: &str,
    source: &str,
    expected: &Result<Shape, usize>,
    label: &str,
) -> Option<grameo::ParseError> {
    let captured = parse_program(program, entry, source, MatchOptions::default());
    let recognized = recognize_program(program, entry, source, MatchOptions::default());
    match (expected, captured, recognized) {
        (Ok(expected), Ok(tree), Ok(recognition)) => {
            let actual = native_shape(&tree);
            if actual != *expected {
                let index = actual
                    .iter()
                    .zip(expected)
                    .position(|(a, b)| a != b)
                    .unwrap_or(actual.len().min(expected.len()));
                panic!(
                    "{label}: capture {index}: native={:?}, reference={:?}; native/reference counts={}/{}",
                    actual.get(index),
                    expected.get(index),
                    actual.len(),
                    expected.len()
                );
            }
            assert_eq!(tree.consumed(), source.len(), "{label}");
            assert_eq!(recognition.consumed, tree.consumed(), "{label}");
            None
        }
        (Err(_), Err(captured), Err(recognized)) => {
            assert_eq!(captured.kind, ParseErrorKind::Syntax, "{label}: {captured}");
            // Native diagnostics identify terminal failures; the reference
            // identifies rule attempts. Measure the difference separately.
            assert!(
                source.is_char_boundary(captured.offset),
                "{label}: diagnostic offset"
            );
            assert!(
                !captured.expected.is_empty(),
                "{label}: missing failure description"
            );
            assert_eq!(recognized, captured, "{label}: recognition diagnostics");
            Some(captured)
        }
        (expected, captured, recognized) => panic!(
            "{label}: {source:?}\nreference={expected:?}\ncaptures={captured:?}\nrecognition={recognized:?}"
        ),
    }
}
// FNV-1a is a reproducibility fingerprint, not a security hash. Length prefixes
// distinguish concatenations with different grammar/path/source boundaries.
fn fingerprint<'a>(parts: impl IntoIterator<Item = &'a str>) -> u64 {
    let mut hash = 0xcbf29ce484222325u64;
    for part in parts {
        for byte in (part.len() as u64)
            .to_le_bytes()
            .into_iter()
            .chain(part.bytes())
        {
            hash = (hash ^ u64::from(byte)).wrapping_mul(0x100000001b3);
        }
    }
    hash
}

fn escaped(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('\t', "\\t")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
}
fn verdict<T>(result: &Result<T, grameo::ParseError>) -> String {
    match result {
        Ok(_) => "accepted".into(),
        Err(error) => format!("{:?}", error.kind),
    }
}
// Full mode visits every matching file, including large files, and writes a row
// immediately so interrupted runs remain inspectable. It deliberately retains
// default matcher limits: resource failures block switching rather than being
// reclassified as equivalent syntax rejection.
fn full_corpus<P: Program, Q: Program>(
    language: &str,
    directory: &str,
    extensions: &[&str],
    entry: &str,
    resolved: &P,
    generated: &Q,
    reference: impl Fn(&str) -> Result<Shape, usize>,
) {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../../tests")
        .join(directory);
    let output = PathBuf::from(
        std::env::var_os("VYBE_PARSER_AUDIT_DIR")
            .expect("full corpus requires VYBE_PARSER_AUDIT_DIR"),
    );
    fs::create_dir_all(&output).unwrap();
    let mut report = fs::File::create(output.join(format!("{language}.tsv"))).unwrap();
    writeln!(report, "label\tbytes\tfingerprint\treference\tresolved_capture\tgenerated_capture\tresolved_recognition\tgenerated_recognition\tstatus\tdiagnostic_offset_difference").unwrap();
    let mut files = Vec::new();
    if let Some(replay) = std::env::var_os("VYBE_PARSER_REPLAY_INPUT") {
        files.push(PathBuf::from(replay));
    } else {
        paths(&root, extensions, &mut files);
    }
    files.sort();
    assert!(!files.is_empty());
    let mut unreadable_report =
        fs::File::create(output.join(format!("{language}.unreadable.tsv"))).unwrap();
    writeln!(unreadable_report, "label\tstatus\tdetails").unwrap();
    let mut cases = 0usize;
    let mut failures = 0usize;
    let mut accepted = 0usize;
    let mut rejected = 0usize;
    let mut resources = 0usize;
    let mut unreadable = 0usize;
    let mut audit = |label: &str, source: &str| {
        cases += 1;
        let expected = reference(source);
        accepted += usize::from(expected.is_ok());
        rejected += usize::from(expected.is_err());
        let captures = [
            parse_program(resolved, entry, source, MatchOptions::default()),
            parse_program(generated, entry, source, MatchOptions::default()),
        ];
        let recognizes = [
            recognize_program(resolved, entry, source, MatchOptions::default()),
            recognize_program(generated, entry, source, MatchOptions::default()),
        ];
        let resource = captures
            .iter()
            .filter_map(|result| result.as_ref().err())
            .chain(recognizes.iter().filter_map(|result| result.as_ref().err()))
            .any(|error| {
                matches!(
                    error.kind,
                    ParseErrorKind::WorkLimit | ParseErrorKind::DepthLimit
                )
            });
        let mut detail = String::new();
        for (backend, capture, recognition) in [
            ("resolved", &captures[0], &recognizes[0]),
            ("generated", &captures[1], &recognizes[1]),
        ] {
            match (&expected, capture, recognition) {
                (Ok(shape), Ok(tree), Ok(recognition)) => {
                    let actual = native_shape(tree);
                    if actual != *shape {
                        let index = actual
                            .iter()
                            .zip(shape)
                            .position(|(a, b)| a != b)
                            .unwrap_or(actual.len().min(shape.len()));
                        detail.push_str(&format!("{backend} capture {index}: native={:?}, reference={:?}; counts={}/{}\n", actual.get(index), shape.get(index), actual.len(), shape.len()));
                    }
                    if tree.consumed() != source.len() || recognition.consumed != tree.consumed() {
                        detail.push_str(&format!(
                            "{backend} consumed: capture={}, recognition={}, source={}\n",
                            tree.consumed(),
                            recognition.consumed,
                            source.len()
                        ));
                    }
                }
                (Err(_), Err(capture), Err(recognition))
                    if capture.kind == ParseErrorKind::Syntax && capture == recognition =>
                {
                    if !source.is_char_boundary(capture.offset) || capture.expected.is_empty() {
                        detail.push_str(&format!(
                            "{backend} invalid syntax diagnostics: {capture:?}\n"
                        ));
                    }
                }
                _ => detail.push_str(&format!(
                    "{backend}: reference={}, capture={}, recognition={}\n",
                    if expected.is_ok() {
                        "accepted"
                    } else {
                        "rejected"
                    },
                    verdict(capture),
                    verdict(recognition)
                )),
            }
        }
        if let (Err(a), Err(b)) = (&captures[0], &captures[1])
            && a != b
        {
            detail.push_str(&format!(
                "native backend diagnostics differ: {a:?} != {b:?}\n"
            ));
        }
        let status = if detail.is_empty() {
            "parity"
        } else if resource {
            "resource_mismatch"
        } else {
            "semantic_mismatch"
        };
        if !detail.is_empty() {
            failures += 1;
            resources += usize::from(resource);
            let stem = format!("{language}-{cases:06}");
            fs::write(output.join(format!("{stem}.source")), source).unwrap();
            fs::write(output.join(format!("{stem}.failure")), format!("label={label}\nentry={entry}\n{detail}\ncapture diagnostics={:?}\nrecognition diagnostics={:?}\n", captures.iter().map(|r| r.as_ref().err()).collect::<Vec<_>>(), recognizes.iter().map(|r| r.as_ref().err()).collect::<Vec<_>>())).unwrap();
            eprintln!("AUDIT_MISMATCH {language} {label}: {status}; reproduction={stem}.source");
        }
        let offset_diff = match (&expected, &captures[0]) {
            (Err(offset), Err(error)) => usize::from(*offset != error.offset),
            _ => 0,
        };
        writeln!(
            report,
            "{}\t{}\t{:016x}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
            escaped(label),
            source.len(),
            fingerprint([label, source]),
            if expected.is_ok() {
                "accepted"
            } else {
                "rejected"
            },
            verdict(&captures[0]),
            verdict(&captures[1]),
            verdict(&recognizes[0]),
            verdict(&recognizes[1]),
            status,
            offset_diff
        )
        .unwrap();
        if cases.is_multiple_of(1000) {
            report.flush().unwrap();
            eprintln!("AUDIT_PROGRESS {language}: {cases} inputs, {failures} mismatches");
        }
    };
    // Match raw input everywhere; the existing two standalone transformations
    // get their own rows and fingerprints, never silently replace original input.
    for path in &files {
        let label = path
            .strip_prefix(&root)
            .unwrap_or(path)
            .display()
            .to_string();
        let bytes = match fs::read(path) {
            Ok(bytes) => bytes,
            Err(error) => {
                unreadable += 1;
                eprintln!("AUDIT_UNREADABLE {language} {label}: {error}");
                writeln!(
                    unreadable_report,
                    "{}\tread_error\t{}",
                    escaped(&label),
                    escaped(&error.to_string())
                )
                .unwrap();
                continue;
            }
        };
        let source = match String::from_utf8(bytes) {
            Ok(source) => source,
            Err(error) => {
                unreadable += 1;
                eprintln!("AUDIT_INVALID_UTF8 {language} {label}: {error}");
                writeln!(
                    unreadable_report,
                    "{}\tinvalid_utf8\t{}",
                    escaped(&label),
                    escaped(&error.to_string())
                )
                .unwrap();
                continue;
            }
        };
        audit(&label, &source);
        let transformed: Option<String> = match language {
            #[cfg(feature = "corpus-scripting")]
            "python" => Some(python_preprocessor::preprocess_indentation(&source)),
            #[cfg(feature = "corpus-business")]
            "fortran" => Some(fortran_preprocessor::rewrite_hollerith(&source)),
            _ => None,
        };
        if let Some(source) = transformed {
            audit(&format!("preprocessed:{label}"), &source);
        }
    }
    report.flush().unwrap();
    eprintln!(
        "AUDIT_SUMMARY {language}: files={} inputs={cases} accepted={accepted} rejected={rejected} mismatches={failures} resource_mismatches={resources} unreadable={unreadable}",
        files.len()
    );
    fs::write(output.join(format!("{language}.summary")), format!("files={}\ninputs={cases}\naccepted={accepted}\nrejected={rejected}\nmismatches={failures}\nresource_mismatches={resources}\nunreadable={unreadable}\n", files.len())).unwrap();
    assert_eq!(
        failures + unreadable,
        0,
        "{language}: full corpus gate failed; inspect {}",
        output.display()
    );
}

macro_rules! language {
    ($name:ident, $feature:literal, $grammar:literal, $entry:ident, $directory:literal, [$($ext:literal),+], [$($seed:literal),+]) => {
        #[cfg(feature = $feature)]
        mod $name {
            use super::*;
            #[derive(pest_derive::Parser)]
            #[grammar = $grammar]
            struct Reference;
            fn reference(source: &str) -> Result<Shape, usize> {
                Reference::parse(Rule::$entry, source).map(reference_shape).map_err(|error| match error.location {
                    pest::error::InputLocation::Pos(offset) | pest::error::InputLocation::Span((offset, _)) => offset,
                })
            }
            #[test]
            fn repository_source_acceptance_capture_and_native_diagnostic_parity() {
                let grammar = grameo::compile(include_str!(concat!("../../../../languages/", stringify!($name), "/src/grammar.pest"))).unwrap();
                if std::env::var_os("VYBE_PARSER_FULL_CORPUS").is_some() {
                    full_corpus(stringify!($name), $directory, &[$($ext),+], stringify!($entry), &grammar,
                        &vybe_parser_generated_tests::all::$name::Parser, reference);
                    return;
                }
                let mut samples = vec![$((format!("seed:{}", $seed), $seed.to_owned())),+];
                for (label, source) in &samples { assert!(reference(source).is_ok(), "{} {label}: seed is rejected", stringify!($name)); }
                samples.extend(corpus($directory, &[$($ext),+]));
                // Corpus includes raw inputs and, for Python/Fortran, inputs transformed
                // with the existing standalone preprocessor helpers. Coordinates
                // are in the matched input; source maps/normalization remain separate.
                samples.push(("empty".into(), String::new()));
                let original = samples.len();
                for index in 0..original {
                    let (label, source) = &samples[index];
                    let boundaries: Vec<_> = source.char_indices().map(|(offset, _)| offset).collect();
                    let prefixes = boundaries.get(boundaries.len() / 2).into_iter().chain(boundaries.last());
                    let additions = prefixes.map(|&end| (format!("{label}:prefix:{end}"), source[..end].to_owned())).collect::<Vec<_>>();
                    let junk = (format!("{label}:trailing-junk"), format!("{source}\0"));
                    samples.extend(additions);
                    samples.push(junk);
                }
                let corpus_fingerprint = fingerprint(std::iter::once(include_str!(concat!("../../../../languages/", stringify!($name), "/src/grammar.pest"))).chain(samples.iter().flat_map(|(label, source)| [label.as_str(), source.as_str()])));
                let mut accepted = 0;
                let mut diagnostic_offsets = 0;
                for (label, source) in &samples {
                    let expected = reference(source);
                    accepted += usize::from(expected.is_ok());
                    let label = format!("{}:{label}", stringify!($name));
                    let resolved = check(&grammar, stringify!($entry), source, &expected, &format!("resolved:{label}"));
                    let generated = check(&vybe_parser_generated_tests::all::$name::Parser, stringify!($entry), source, &expected, &format!("generated:{label}"));
                    assert_eq!(resolved, generated, "{label}: backend diagnostics");
                    if let (Err(reference_offset), Some(native)) = (&expected, resolved) {
                        diagnostic_offsets += usize::from(*reference_offset != native.offset);
                    }
                }
                eprintln!("{}: {} inputs, {} accepted, {} rejected, {} reference/native diagnostic offset differences; fingerprint={:016x}; resolved/generated capture+recognition parity", stringify!($name), samples.len(), accepted, samples.len()-accepted, diagnostic_offsets, corpus_fingerprint);
            }
        }
    }
}

language!(
    js,
    "corpus-scripting",
    "../../../languages/js/src/grammar.pest",
    program,
    "js",
    ["js"],
    [
        "let x = 1; x += 2;",
        "function f(x) { return x + 1; } f(2);"
    ]
);

language!(
    php,
    "corpus-scripting",
    "../../../languages/php/src/grammar.pest",
    program,
    "php",
    ["php"],
    [
        "<?php $x = 1; echo $x;",
        "function f($x) { return $x + 1; }"
    ]
);

language!(
    python,
    "corpus-scripting",
    "../../../languages/python/src/grammar.pest",
    program,
    "python",
    ["py"],
    ["x = 1\nprint(x)\n", "if True:\n⇥print(1)\n⇤"]
);

language!(
    c,
    "corpus-c-family",
    "../../../languages/c/src/grammar.pest",
    program,
    "c",
    ["c"],
    ["int main() { return 0; }", "int x = 1;"]
);

language!(
    csharp,
    "corpus-c-family",
    "../../../languages/csharp/src/grammar.pest",
    program,
    "csharp",
    ["cs"],
    ["class Demo { static void Main() { } }", "int x = 1;"]
);

language!(
    java,
    "corpus-c-family",
    "../../../languages/java/src/grammar.pest",
    program,
    "java",
    ["java"],
    [
        "class Demo { public static void main(String[] args) { } }",
        "class X { int x = 1; }"
    ]
);

language!(
    dart,
    "corpus-jvm-go-dart",
    "../../../languages/dart/src/grammar.pest",
    program,
    "dart",
    ["dart"],
    ["void main() { print(1); }", "int x = 1;"]
);

language!(
    go,
    "corpus-jvm-go-dart",
    "../../../languages/go/src/grammar.pest",
    program,
    "go",
    ["go"],
    [
        "package main\nfunc main() { println(1) }\n",
        "package main\nvar x int = 1\n"
    ]
);

language!(
    kotlin,
    "corpus-jvm-go-dart",
    "../../../languages/kotlin/src/grammar.pest",
    program,
    "kotlin",
    ["kt"],
    [
        "val x = 1\nprintln(x)\n",
        "fun f(x: Int): Int { return x + 1 }\n"
    ]
);

language!(
    lua,
    "corpus-dynamic",
    "../../../languages/lua/src/grammar.pest",
    chunk,
    "lua",
    ["lua"],
    ["local x = 1; return x", "for i=1,3 do print(i) end"]
);

language!(
    ruby,
    "corpus-dynamic",
    "../../../languages/ruby/src/grammar.pest",
    program,
    "ruby",
    ["rb"],
    ["x = 1\nputs x\n", "def f(x)\nreturn x + 1\nend\n"]
);

language!(
    powershell,
    "corpus-dynamic",
    "../../../languages/powershell/src/grammar.pest",
    program,
    "powershell",
    ["ps1"],
    ["$x = 1\nWrite-Output $x\n", "function f { return 1 }\nf\n"]
);

language!(
    cobol,
    "corpus-business",
    "../../../languages/cobol/src/grammar.pest",
    program,
    "cobol",
    ["cob"],
    [
        "IDENTIFICATION DIVISION. PROGRAM-ID. DEMO. PROCEDURE DIVISION. DISPLAY 1. STOP RUN.",
        "IDENTIFICATION DIVISION. PROGRAM-ID. DEMO."
    ]
);

language!(
    fortran,
    "corpus-business",
    "../../../languages/fortran/src/grammar.pest",
    program,
    "fortran",
    ["f90", "f"],
    [
        "program demo\nprint *, 1\nend program demo\n",
        "integer :: x\nx = 1\n"
    ]
);

language!(
    pascal,
    "corpus-business",
    "../../../languages/pascal/src/grammar.pest",
    program,
    "pascal",
    ["pas"],
    ["program Demo; begin writeln(1); end.", "begin end."]
);

language!(
    vb,
    "corpus-vb-shell-wast",
    "../../../languages/vb/src/grammar.pest",
    program,
    "vb",
    ["vb"],
    [
        "Dim x As Integer = 1\nConsole.WriteLine(x)\n",
        "Sub Main()\nEnd Sub\n"
    ]
);

language!(
    washm,
    "corpus-vb-shell-wast",
    "../../../languages/washm/src/grammar.pest",
    program,
    "shell",
    ["sh"],
    ["echo hello\n", "for x in a b; do echo \"$x\"; done\n"]
);

language!(
    wast,
    "corpus-vb-shell-wast",
    "../../../languages/wast/src/grammar.pest",
    program,
    "wast",
    ["wat", "wast"],
    [
        "(module)",
        "(module (func (export \"f\") (result i32) i32.const 1))"
    ]
);
