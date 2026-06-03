use super::helpers::run_prints;

// ── Exception $previous / chaining ───────────────────────────

#[test]
fn exception_previous_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    try { throw new RuntimeException('original'); }
    catch (RuntimeException $e) { throw new LogicException('wrapped', 0, $e); }
} catch (LogicException $e) {
    echo $e->getMessage() . ':' . $e->getPrevious()->getMessage();
}
"#
        ),
        vec!["wrapped:original"]
    );
}
#[test]
fn exception_previous_is_null_when_not_chained() {
    assert_eq!(
        run_prints(
            r#"<?php
try { throw new Exception('plain'); }
catch (Exception $e) { echo $e->getPrevious() === null ? 'null' : 'has prev'; }
"#
        ),
        vec!["null"]
    );
}
#[test]
fn exception_chain_three_levels() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    try {
        try { throw new Exception('root'); }
        catch (Exception $e) { throw new RuntimeException('mid', 0, $e); }
    } catch (RuntimeException $e) { throw new LogicException('top', 0, $e); }
} catch (LogicException $e) {
    $chain = [];
    $cur = $e;
    while ($cur !== null) { $chain[] = $cur->getMessage(); $cur = $cur->getPrevious(); }
    echo implode('>', $chain);
}
"#
        ),
        vec!["top>mid>root"]
    );
}

// ── Exception properties ──────────────────────────────────────

#[test]
fn exception_code_property() {
    assert_eq!(
        run_prints(
            r#"<?php
try { throw new Exception('msg', 404); }
catch (Exception $e) { echo $e->getCode() . ':' . $e->getMessage(); }
"#
        ),
        vec!["404:msg"]
    );
}
#[test]
fn exception_file_and_line_set() {
    assert_eq!(
        run_prints(
            r#"<?php
try { throw new Exception('x'); }
catch (Exception $e) { echo $e->getFile() !== '' ? 'has_file' : 'no_file'; }
"#
        ),
        vec!["has_file"]
    );
}
#[test]
fn exception_to_string() {
    assert_eq!(
        run_prints(
            r#"<?php
try { throw new RuntimeException('boom', 500); }
catch (RuntimeException $e) {
    echo str_contains((string)$e, 'RuntimeException') ? 'yes' : 'no';
}
"#
        ),
        vec!["yes"]
    );
}

// ── Custom exception classes ──────────────────────────────────

#[test]
fn custom_exception_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class DomainException2 extends RuntimeException {
    public function __construct(string $msg, private string $domain) {
        parent::__construct($msg);
    }
    public function getDomain(): string { return $this->domain; }
}
try { throw new DomainException2('err', 'payments'); }
catch (DomainException2 $e) { echo $e->getMessage() . ':' . $e->getDomain(); }
"#
        ),
        vec!["err:payments"]
    );
}
#[test]
fn custom_exception_hierarchy() {
    assert_eq!(
        run_prints(
            r#"<?php
class AppException extends RuntimeException {}
class DbException extends AppException {}
try { throw new DbException('conn failed'); }
catch (AppException $e) { echo 'caught:' . get_class($e); }
"#
        ),
        vec!["caught:DbException"]
    );
}
#[test]
fn exception_message_after_rethrow() {
    assert_eq!(
        run_prints(
            r#"<?php
function risky(): void { throw new RuntimeException('initial'); }
try {
    try { risky(); }
    catch (RuntimeException $e) { throw new RuntimeException('rethrown', 0, $e); }
} catch (RuntimeException $e) {
    echo $e->getMessage() . '/' . $e->getPrevious()->getMessage();
}
"#
        ),
        vec!["rethrown/initial"]
    );
}

// ── PHP error hierarchy ───────────────────────────────────────

#[test]
fn type_error_caught_by_error() {
    assert_eq!(
        run_prints(
            r#"<?php
function add(int $a, int $b): int { return $a + $b; }
try { add('x', 1); }
catch (TypeError $e) { echo 'TypeError'; }
"#
        ),
        vec!["TypeError"]
    );
}
#[test]
fn value_error_from_array_function() {
    assert_eq!(
        run_prints(
            r#"<?php
try { array_chunk([], 0); }
catch (ValueError $e) { echo 'ValueError'; }
"#
        ),
        vec!["ValueError"]
    );
}
#[test]
fn division_by_zero_error() {
    assert_eq!(
        run_prints(
            r#"<?php
try { $r = intdiv(5, 0); }
catch (\DivisionByZeroError $e) { echo 'DivByZero'; }
"#
        ),
        vec!["DivByZero"]
    );
}
#[test]
fn error_is_not_exception() {
    assert_eq!(
        run_prints(
            r#"<?php
try { throw new Error('low level'); }
catch (Exception $e) { echo 'exception'; }
catch (Error $e) { echo 'error'; }
"#
        ),
        vec!["error"]
    );
}
#[test]
fn throwable_catches_both() {
    assert_eq!(
        run_prints(
            r#"<?php
function catchAll(\Throwable $t): string { return get_class($t); }
echo catchAll(new Exception('e')) . ',' . catchAll(new Error('err'));
"#
        ),
        vec!["Exception,Error"]
    );
}

// ── Finally and exception propagation ────────────────────────

#[test]
fn finally_runs_even_when_rethrown() {
    assert_eq!(
        run_prints(
            r#"<?php
$log = [];
try {
    try {
        throw new Exception('e');
    } finally {
        $log[] = 'inner_finally';
    }
} catch (Exception $e) {
    $log[] = 'caught';
}
echo implode(',', $log);
"#
        ),
        vec!["inner_finally,caught"]
    );
}
#[test]
fn finally_runs_on_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function withFinally(): string {
    try { return 'from_try'; }
    finally { echo 'cleanup,'; }
}
echo withFinally();
"#
        ),
        vec!["cleanup,from_try"]
    );
}
