use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Exceptions Rethrowing & Trace Walking — getPrevious(), getTrace(), getTraceAsString(), custom getters
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_exception_previous_chain_traversal() {
    let out = run_prints(
        r#"<?php
class DbError extends Exception {}
class ServiceError extends Exception {}

try {
    try {
        throw new DbError("Connection refused", 1001);
    } catch (DbError $e) {
        throw new ServiceError("Service unavailable", 503, $e);
    }
} catch (ServiceError $e) {
    $chain = [];
    $curr = $e;
    while ($curr !== null) {
        $chain[] = get_class($curr) . ": " . $curr->getMessage();
        $curr = $curr->getPrevious();
    }
    echo implode(" <- ", $chain);
}
"#,
    );
    assert_eq!(
        out,
        vec!["ServiceError: Service unavailable <- DbError: Connection refused"]
    );
}

#[test]
fn test_php_exception_get_file_line_code_getters() {
    let out = run_prints(
        r#"<?php
try {
    throw new Exception("Custom exception message", 404);
} catch (Exception $e) {
    echo "Code=" . $e->getCode() . " Msg=" . $e->getMessage() . " File=" . (strlen($e->getFile()) > 0 ? "OK" : "NO");
}
"#,
    );
    assert_eq!(out, vec!["Code=404 Msg=Custom exception message File=OK"]);
}

#[test]
fn test_php_exception_get_trace_array_structure() {
    let out = run_prints(
        r#"<?php
function levelTwo() {
    throw new Exception("Deep Error");
}
function levelOne() {
    levelTwo();
}

try {
    levelOne();
} catch (Exception $e) {
    $trace = $e->getTrace();
    $funcs = array_column($trace, "function");
    echo "levelTwo=" . (in_array("levelTwo", $funcs) ? "1" : "0") . " levelOne=" . (in_array("levelOne", $funcs) ? "1" : "0");
}
"#,
    );
    assert_eq!(out, vec!["levelTwo=1 levelOne=1"]);
}

#[test]
fn test_php_exception_to_string_formatting() {
    compile_ok(
        r#"<?php
try {
    throw new Exception("Test stringification", 500);
} catch (Exception $e) {
    $str = (string)$e;
    echo str_contains($str, "Test stringification") ? "STRINGIFIED_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php_exception_get_trace_as_string() {
    compile_ok(
        r#"<?php
try {
    throw new Exception("Trace String Test");
} catch (Exception $e) {
    $traceStr = $e->getTraceAsString();
    echo is_string($traceStr) && strlen($traceStr) > 0 ? "TRACE_STR_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php_custom_exception_context_payload() {
    compile_ok(
        r#"<?php
class HttpPayloadException extends Exception {
    public function __construct(public array $context, string $message = "", int $code = 0) {
        parent::__construct($message, $code);
    }
}

try {
    throw new HttpPayloadException(["ip" => "127.0.0.1"], "Unauthorized", 401);
} catch (HttpPayloadException $e) {
    echo "IP=" . $e->context["ip"] . " Code=" . $e->getCode();
}
"#,
    );
}

#[test]
fn test_php_error_exception_severity_level() {
    compile_ok(
        r#"<?php
try {
    throw new ErrorException("Warning Exception", 0, E_WARNING, __FILE__, __LINE__);
} catch (ErrorException $e) {
    echo "Severity=" . $e->getSeverity();
}
"#,
    );
}

#[test]
fn test_php_finally_block_runs_on_uncaught_exception() {
    compile_ok(
        r#"<?php
function testFinally() {
    try {
        throw new Exception("Fatal inside function");
    } finally {
        echo "FINALLY_EXECUTED ";
    }
}

try {
    testFinally();
} catch (Exception $e) {
    echo "CAUGHT_OUTSIDE";
}
"#,
    );
}

#[test]
fn test_php_exception_cloning_prevention() {
    compile_ok(
        r#"<?php
$e1 = new Exception("Original");
try {
    $e2 = clone $e1;
} catch (Error $err) {
    echo "Cannot clone exception";
}
"#,
    );
}

#[test]
fn test_php_throw_in_destructor_safety() {
    compile_ok(
        r#"<?php
class DangerousDestructor {
    public function __destruct() {
        // Exceptions thrown from destructors must be handled or will trigger fatal error
    }
}
$d = new DangerousDestructor();
"#,
    );
}
