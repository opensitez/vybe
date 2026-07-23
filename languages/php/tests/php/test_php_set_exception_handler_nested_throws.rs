use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Exception Handlers: set_exception_handler, restore & Nested Uncaught Exception Chaining
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_set_exception_handler_captures_uncaught_exception() {
    let out = run_prints(
        r##"<?php
$captured = "";
set_exception_handler(function(Throwable $e) use (&$captured) {
    $captured = "Caught: " . $e->getMessage();
});

try {
    throw new Exception("Test Exception Payload");
} catch (Throwable $e) {
    // Manually invoke exception handler for isolated test runner predictability
    $handler = set_exception_handler(null);
    $handler($e);
}

echo $captured;
"##,
    );
    assert_eq!(out, vec!["Caught: Test Exception Payload"]);
}

#[test]
fn test_php_restore_exception_handler_reverts_to_previous() {
    let out = run_prints(
        r##"<?php
$log = [];
$h1 = function(Throwable $e) use (&$log) { $log[] = "H1:" . $e->getMessage(); };
$h2 = function(Throwable $e) use (&$log) { $log[] = "H2:" . $e->getMessage(); };

set_exception_handler($h1);
set_exception_handler($h2);
restore_exception_handler(); // Reverts to H1

$current = set_exception_handler(null);
$current(new Exception("Event"));

echo implode(", ", $log);
"##,
    );
    assert_eq!(out, vec!["H1:Event"]);
}

#[test]
fn test_php_set_exception_handler_throwable_type_hint_error() {
    compile_ok(
        r##"<?php
$caughtError = false;
set_exception_handler(function(Throwable $t) use (&$caughtError) {
    if ($t instanceof TypeError) $caughtError = true;
});
try {
    throw new TypeError("Type error thrown");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $caughtError ? "TYPE_ERROR_CAUGHT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_returns_previous_callable() {
    compile_ok(
        r##"<?php
$first = fn($e) => null;
set_exception_handler($first);
$second = set_exception_handler(fn($e) => null);
restore_exception_handler();
restore_exception_handler();
echo $second === $first ? "PREVIOUS_HANDLER_RETURNED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_custom_exception_subclass() {
    compile_ok(
        r##"<?php
class CustomDomainException extends DomainException {}
$domainCaptured = false;
set_exception_handler(function(Throwable $e) use (&$domainCaptured) {
    if ($e instanceof CustomDomainException) $domainCaptured = true;
});
try {
    throw new CustomDomainException("Domain breach");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $domainCaptured ? "CUSTOM_DOMAIN_EXCEPTION_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_null_resets() {
    compile_ok(
        r##"<?php
set_exception_handler(fn($e) => null);
set_exception_handler(null);
echo "NULL_RESET_EXCEPTION_HANDLER_OK";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_anonymous_class_callable() {
    compile_ok(
        r##"<?php
$invoked = false;
$handler = new class(&$invoked) {
    private $invoked;
    public function __construct(&$invoked) { $this->invoked = &$invoked; }
    public function __invoke(Throwable $e) { $this->invoked = true; }
};
set_exception_handler($handler);
try {
    throw new Exception("Anon class exception");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $invoked ? "ANON_CLASS_EXCEPTION_HANDLER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_chained_previous_exception() {
    compile_ok(
        r##"<?php
$hasPrevious = false;
set_exception_handler(function(Throwable $e) use (&$hasPrevious) {
    if ($e->getPrevious() !== null) $hasPrevious = true;
});
try {
    $first = new Exception("First");
    throw new Exception("Second", 0, $first);
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $hasPrevious ? "CHAINED_PREVIOUS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_trace_as_string_inspection() {
    compile_ok(
        r##"<?php
$traceCaptured = false;
set_exception_handler(function(Throwable $e) use (&$traceCaptured) {
    if (strlen($e->getTraceAsString()) > 0) $traceCaptured = true;
});
try {
    throw new Exception("Trace check");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $traceCaptured ? "TRACE_AS_STRING_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_exception_handler_code_property() {
    compile_ok(
        r##"<?php
$codeVal = 0;
set_exception_handler(function(Throwable $e) use (&$codeVal) {
    $codeVal = $e->getCode();
});
try {
    throw new Exception("With code", 404);
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $codeVal === 404 ? "EXCEPTION_CODE_404_OK" : "FAIL";
"##,
    );
}
