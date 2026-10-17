use grameo::source_units::PreparationContext;

fn assert_module_parity(source: &str, module: &vybe_ast::Module) {
    assert_eq!(format!("{module:?}"), format!("{:?}", vybe_language_php::parse(source).unwrap()), "{source}");
    assert_eq!(format!("{module:?}"), format!("{:?}", vybe_language_php::parse_with_peg(source).unwrap()), "{source}");
}

#[test]
fn rejected_pratt_islands_retain_each_owned_include_receipt() {
    for (source, targets) in [
        ("<?php return 1 <=> (include 'a');", vec!["'a'"]),
        ("<?php return $a <=> foo(include 'a');", vec!["'a'"]),
        ("<?php return foo(include 'a') <=> bar(require_once 'b');", vec!["'a'", "'b'"]),
        ("<?php return (include 'a') <=> (require 'b') + (include_once 'c');", vec!["'a'", "'b'", "'c'"]),
        ("<b>before</b><?php return 1 <=> (include (require 'a')); ?>after", vec!["'a'", "(require 'a')"]),
        ("<?php return (1 <=> (include 'a')) && (include_once 'b');", vec!["'a'", "'b'"]),
    ] {
        let (module, requests) = vybe_language_php::parse_with_source_units(source).unwrap();
        assert_module_parity(source, &module);
        assert_eq!(requests.iter().map(|request| request.target.trim()).collect::<Vec<_>>(), targets, "{source}");
        assert!(requests.iter().all(|request| request.preparation == PreparationContext::Immediate));
    }
}

#[test]
fn nested_pratt_replay_preserves_deferred_scope_and_one_committed_batch() {
    let source = "<?php function f() { return 1 <=> (include 'function.php'); } class C { function m() { return $a <=> foo(include_once 'method.php'); } } $f = function () { return 1 <=> (require 'closure.php'); }; $g = fn() => 1 <=> (require_once 'arrow.php'); include 'tail.php';";
    let (expected, requests) = vybe_language_php::parse_with_source_units(source).unwrap();
    assert_module_parity(source, &expected);
    assert_eq!(requests.iter().map(|request| request.preparation).collect::<Vec<_>>(), [
        PreparationContext::Deferred, PreparationContext::Deferred,
        PreparationContext::Deferred, PreparationContext::Deferred,
        PreparationContext::Immediate,
    ]);
    let mut batches = Vec::new();
    let module = vybe_language_php::parse_with_source_unit_preparation(source, &mut |batch| batches.push(batch)).unwrap();
    assert_eq!(format!("{module:?}"), format!("{expected:?}"));
    assert_eq!(batches, [requests]);
}

#[test]
fn later_parse_failure_rolls_back_replayed_receipts_before_publication() {
    for source in [
        "<?php include 'first.php'; return 1 <=> (include 'lost.php') +;",
        "<?php function f() { return $a <=> foo(require 'lost.php'); } function broken(",
        "<p>before</p><?php return 1 <=> (include (require 'lost.php')) +; ?>after",
    ] {
        let mut publications = 0;
        assert!(vybe_language_php::parse_with_source_unit_preparation(source, &mut |_| publications += 1).is_err(), "{source}");
        assert_eq!(publications, 0, "{source}");
        assert!(vybe_language_php::parse_with_source_units(source).is_err(), "{source}");
    }
}

/// The WordPress checkout is a local application fixture, not part of the
/// crate's packaged tests. Run explicitly with --ignored when it is available.
#[test]
#[ignore = "requires the local examples/webroot/wordpress application fixture"]
fn wordpress_modules_keep_exact_ast_with_committed_receipts() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/webroot/wordpress");
    for name in ["wp-includes/load.php", "wp-includes/functions.php", "wp-settings.php"] {
        let source = std::fs::read_to_string(root.join(name)).unwrap();
        assert!(source.len() <= 400_000, "bounded fixture unexpectedly grew: {name}");
        let expected = vybe_language_php::parse(&source).unwrap();
        let mut batches = 0;
        let mut requests = 0;
        let actual = vybe_language_php::parse_with_source_unit_preparation(&source, &mut |owned| {
            batches += 1;
            requests += owned.len();
        }).unwrap();
        assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "complete Module diverged: {name}");
        assert_eq!(batches, 1, "grammar commitment must publish once: {name}");
        assert!(requests > 0, "fixture should exercise source intents: {name}");
    }
}
