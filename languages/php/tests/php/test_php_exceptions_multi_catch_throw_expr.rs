use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Exceptions, Multi-Catch & Throw Expressions — Throwable, Exception, Error, multi-catch (A | B), throw expr, chaining
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_multi_catch_exception_handling() {
    let out = run_prints(
        r#"<?php
class InvalidUserException extends Exception {}
class DatabaseException extends Exception {}

function process($type) {
    try {
        if ($type === 1) throw new InvalidUserException("User invalid");
        if ($type === 2) throw new DatabaseException("DB error");
    } catch (InvalidUserException | DatabaseException $e) {
        echo "CAUGHT: " . $e->getMessage();
    }
}

process(2);
"#,
    );
    assert_eq!(out, vec!["CAUGHT: DB error"]);
}

#[test]
fn test_php80_throw_as_expression_in_null_coalescing() {
    let out = run_prints(
        r#"<?php
function getHost(array $config) {
    return $config["host"] ?? throw new InvalidArgumentException("Missing host");
}

try {
    getHost([]);
} catch (InvalidArgumentException $e) {
    echo "EXPR_THROWN: " . $e->getMessage();
}
"#,
    );
    assert_eq!(out, vec!["EXPR_THROWN: Missing host"]);
}

#[test]
fn test_php_exception_chaining_previous() {
    let out = run_prints(
        r#"<?php
try {
    try {
        throw new RuntimeException("Low level I/O fail", 101);
    } catch (RuntimeException $e) {
        throw new LogicException("High level processing error", 500, $e);
    }
} catch (LogicException $e) {
    echo $e->getMessage() . " -> " . $e->getPrevious()->getMessage();
}
"#,
    );
    assert_eq!(
        out,
        vec!["High level processing error -> Low level I/O fail"]
    );
}

#[test]
fn test_php_finally_block_execution_always() {
    let out = run_prints(
        r#"<?php
$log = [];
try {
    $log[] = "try";
    throw new Exception("fail");
} catch (Exception $e) {
    $log[] = "catch";
} finally {
    $log[] = "finally";
}
echo implode("-", $log);
"#,
    );
    assert_eq!(out, vec!["try-catch-finally"]);
}

#[test]
fn test_php_throw_expression_in_arrow_function() {
    compile_ok(
        r#"<?php
$validate = fn($val) => $val > 0 ? $val : throw new InvalidArgumentException("Must be positive");
echo $validate(10);
"#,
    );
}

#[test]
fn test_php_throw_expression_in_match_arm() {
    compile_ok(
        r#"<?php
$action = "invalid";
$res = match ($action) {
    "run" => "running",
    "stop" => "stopped",
    default => throw new DomainException("Invalid action: $action"),
};
"#,
    );
}

#[test]
fn test_php_throwable_interface_polymorphism() {
    compile_ok(
        r#"<?php
function handleAny(Throwable $t) {
    echo "Throwable: " . $t->getMessage() . " at " . $t->getFile() . ":" . $t->getLine();
}

try {
    throw new TypeError("Type mismatch");
} catch (Throwable $t) {
    handleAny($t);
}
"#,
    );
}

#[test]
fn test_php_custom_exception_properties_and_methods() {
    compile_ok(
        r#"<?php
class ValidationException extends Exception {
    public function __construct(public array $errors, string $msg = "Validation failed") {
        parent::__construct($msg);
    }
}

try {
    throw new ValidationException(["email" => "Required", "age" => "Must be > 18"]);
} catch (ValidationException $e) {
    echo implode(", ", $e->errors);
}
"#,
    );
}

#[test]
fn test_php_rethrowing_caught_exception() {
    compile_ok(
        r#"<?php
function inner() {
    throw new Exception("Inner fail");
}

function outer() {
    try {
        inner();
    } catch (Exception $e) {
        // Log & rethrow
        throw $e;
    }
}

try {
    outer();
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_error_exception_conversion() {
    compile_ok(
        r#"<?php
set_error_handler(function($severity, $message, $file, $line) {
    if (!(error_reporting() & $severity)) {
        return;
    }
    throw new ErrorException($message, 0, $severity, $file, $line);
});

try {
    // trigger user notice or warning
    trigger_error("User warning", E_USER_WARNING);
} catch (ErrorException $e) {
    echo "Converted error to exception: " . $e->getMessage();
} finally {
    restore_error_handler();
}
"#,
    );
}
