use super::helpers::{compile_ok, run_prints};

#[test]
fn try_catch() {
    compile_ok("<?php try { throw new Exception('oops'); } catch (Exception $e) { echo $e; }");
}
#[test]
fn try_finally() {
    compile_ok("<?php try { echo 'try'; } finally { echo 'finally'; }");
}
#[test]
fn try_catch_finally() {
    compile_ok(
        "<?php try { echo 'try'; } catch (Exception $e) { echo 'catch'; } finally { echo 'finally'; }",
    );
}
#[test]
fn throw_expr() {
    compile_ok("<?php function fail() { throw new Exception('fail'); }");
}
#[test]
fn catch_no_var() {
    compile_ok("<?php try { throw new Exception('x'); } catch (Exception) { echo 'caught'; }");
}

#[test]
fn try_catch_captures_exception_message_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    throw new Exception('bad');
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ),
        vec!["bad"]
    );
}

#[test]
fn finally_always_runs_on_success_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 'start|';
try {
    echo 'ok';
} finally {
    echo '|finally';
}
echo '|done';
"#,
        ),
        vec!["start|ok|finally|done"]
    );
}

#[test]
fn finally_always_runs_on_exception_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    try {
        throw new Exception('inner');
    } finally {
        echo 'inner-finally|';
    }
} catch (Exception $e) {
    echo $e->getMessage() . '|';
}
echo 'done';
"#,
        ),
        vec!["inner-finally|inner|done"]
    );
}

#[test]
fn exception_type_filter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    throw new RuntimeException('runtime');
} catch (InvalidArgumentException $e) {
    echo 'arg';
} catch (RuntimeException $e) {
    echo $e->getMessage();
} catch (Exception $e) {
    echo 'base';
}
"#,
        ),
        vec!["runtime"]
    );
}

#[test]
fn custom_exception_class_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class DomainFailure extends Exception {}
try {
    throw new DomainFailure('domain');
} catch (DomainFailure $e) {
    echo $e->getCode();
    echo '|';
    echo $e->getMessage();
}
"#,
        ),
        vec!["0|domain"]
    );
}

#[test]
fn nested_try_catch_finally_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function evalValue(int $v): int {
    try {
        if ($v < 0) {
            throw new Exception('negative');
        }
        return $v * 2;
    } catch (Exception $e) {
        return 0;
    } finally {
        return $v + 1;
    }
}
echo evalValue(-1) . '|' . evalValue(3);
"#,
        ),
        vec!["0|4"]
    );
}
