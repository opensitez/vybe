use grameo::source_units::{Missing, Phase, Repetition, Scope, PreparationContext};

#[test]
fn include_receipts_preserve_common_ast_and_each_policy() {
    let source = "<?php function load($basedir) { include $basedir . '/a.php'; include_once ('b.php'); require 'c.php'; require_once $filename; }";
    let (module, requests) = vybe_language_php::parse_with_source_units(source).unwrap();
    assert_eq!(format!("{module:?}"), format!("{:?}", vybe_language_php::parse(source).unwrap()));
    assert_eq!(format!("{module:?}"), format!("{:?}", vybe_language_php::parse_with_peg(source).unwrap()));
    assert_eq!(requests.len(), 4);
    for (request, (repetition, missing)) in requests.iter().zip([
        (Repetition::Always, Missing::Warning), (Repetition::Once, Missing::Warning),
        (Repetition::Always, Missing::Error), (Repetition::Once, Missing::Error),
    ]) {
        assert_eq!(request.loader, "source");
        assert_eq!(request.policy.phase, Phase::Execute);
        assert_eq!(request.policy.scope, Scope::Caller);
        assert_eq!(request.policy.repetition, repetition);
        assert_eq!(request.policy.missing, missing);
        assert_eq!(request.preparation, PreparationContext::Deferred);
    }
    assert_eq!(requests[0].target.trim(), "$basedir . '/a.php'");
    assert_eq!(requests[3].target.trim(), "$filename");
}

#[test]
fn normalized_receipts_do_not_index_raw_html_and_nested_loads_stay_unevaluated() {
    let source = "<h1>before</h1><?php include (include 'path.php'); ?>after";
    let (module, requests) = vybe_language_php::parse_with_source_units(source).unwrap();
    assert_eq!(format!("{module:?}"), format!("{:?}", vybe_language_php::parse(source).unwrap()));
    assert_eq!(requests.len(), 2);
    assert_eq!(requests[0].target.trim(), "'path.php'");
    assert_eq!(requests[1].target.trim(), "(include 'path.php')");
    assert!(vybe_language_php::parse_with_source_units("<?php include (").is_err());
}

#[test]
fn preparation_context_retains_loads_without_speculating_deferred_bodies() {
    let source="<?php include 'first.php'; function f() { include 'function.php'; } class C { function m() { require 'method.php'; } } $f = function () { include 'closure.php'; }; $g = fn() => require_once 'arrow.php'; include 'last.php';";
    let (module,requests)=vybe_language_php::parse_with_source_units(source).unwrap();
    assert_eq!(format!("{module:?}"),format!("{:?}",vybe_language_php::parse(source).unwrap()));
    assert_eq!(requests.iter().map(|request|request.preparation).collect::<Vec<_>>(),[
        PreparationContext::Immediate,PreparationContext::Deferred,PreparationContext::Deferred,
        PreparationContext::Deferred,PreparationContext::Deferred,PreparationContext::Immediate,
    ]);
}

#[test]
fn committed_preparation_transfers_receipts_once_and_syntax_errors_publish_none() {
    let source="<?php require_once 'load.php'; function f() { include $path; }";
    let mut published=Vec::new();
    let module=vybe_language_php::parse_with_source_unit_preparation(source,&mut |requests|published.push(requests)).unwrap();
    let (expected,requests)=vybe_language_php::parse_with_source_units(source).unwrap();
    assert_eq!(format!("{module:?}"),format!("{expected:?}"));
    assert_eq!(published,[requests]);
    let mut failed_calls=0;
    assert!(vybe_language_php::parse_with_source_unit_preparation("<?php include 'x.php'; function broken(",&mut |_|failed_calls+=1).is_err());
    assert_eq!(failed_calls,0);
}

#[test]
fn no_load_receipt_route_keeps_contextual_pratt_and_scope_directives() {
    for source in [
        "<?php function f($x) { global $shared; return g($x) + $shared; }",
        "<?php class C { function f($x) { return $this->g($x) | h($x); } }",
        "<?php $x = f(2) && g(3); $y = $x ?? h(4);",
        "<?php $text = 'require'; /* INCLUDE */ $x = f(2) + g(3);",
    ] {
        let mut published=Vec::new();
        let module=vybe_language_php::parse_with_source_unit_preparation(source,&mut |requests|published.push(requests)).unwrap();
        assert_eq!(format!("{module:?}"),format!("{:?}",vybe_language_php::parse(source).unwrap()),"{source}");
        assert_eq!(format!("{module:?}"),format!("{:?}",vybe_language_php::parse_with_peg(source).unwrap()),"{source}");
        assert_eq!(published,[Vec::new()],"{source}");
    }
}
