use super::helpers::run_prints;

#[test]
fn test_reflection_generator_get_executing_file() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen() {
    yield 1;
}
$g = gen();
$g->current();
$rg = new ReflectionGenerator($g);
echo is_string($rg->getExecutingFile()) ? 'file_ok' : 'err', "\n";
"#
        ),
        vec!["file_ok"]
    );
}

#[test]
fn test_reflection_generator_get_trace() {
    assert_eq!(
        run_prints(
            r#"<?php
function worker() {
    yield 'step1';
}
$g = worker();
$g->current();
$rg = new ReflectionGenerator($g);
$trace = $rg->getTrace();
echo is_array($trace) ? 'trace_array_ok' : 'err', "\n";
"#
        ),
        vec!["trace_array_ok"]
    );
}
