use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: SPL Exception Hierarchy — InvalidArgumentException, LengthException, OutOfRangeException, DomainException, BadMethodCallException
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_logic_exception_subclasses() {
    let out = run_prints(
        r#"<?php
function checkAge(int $age) {
    if ($age < 0) throw new InvalidArgumentException("Age negative");
    if ($age > 150) throw new OutOfRangeException("Age out of bounds");
}

try {
    checkAge(-5);
} catch (LogicException $e) { // InvalidArgumentException extends LogicException
    echo "LogicException: " . get_class($e) . " -> " . $e->getMessage();
}
"#,
    );
    assert_eq!(
        out,
        vec!["LogicException: InvalidArgumentException -> Age negative"]
    );
}

#[test]
fn test_php_spl_runtime_exception_subclasses() {
    let out = run_prints(
        r#"<?php
function processBuffer(string $data) {
    if (strlen($data) === 0) throw new UnderflowException("Buffer empty");
    if (strlen($data) > 100) throw new OverflowException("Buffer full");
}

try {
    processBuffer("");
} catch (RuntimeException $e) { // UnderflowException extends RuntimeException
    echo "RuntimeException: " . get_class($e) . " -> " . $e->getMessage();
}
"#,
    );
    assert_eq!(
        out,
        vec!["RuntimeException: UnderflowException -> Buffer empty"]
    );
}

#[test]
fn test_php_bad_method_call_exception_handling() {
    let out = run_prints(
        r#"<?php
class GuardedProxy {
    public function __call(string $name, array $args) {
        throw new BadMethodCallException("Method '$name' not supported on proxy");
    }
}

$proxy = new GuardedProxy();
try {
    $proxy->undefinedMethod();
} catch (BadFunctionCallException $e) { // BadMethodCallException extends BadFunctionCallException
    echo "BadCall: " . $e->getMessage();
}
"#,
    );
    assert_eq!(
        out,
        vec!["BadCall: Method 'undefinedMethod' not supported on proxy"]
    );
}

#[test]
fn test_php_domain_exception_business_logic_violation() {
    compile_ok(
        r#"<?php
function calculateDiscount(float $price, float $discount) {
    if ($discount > $price) {
        throw new DomainException("Discount exceeds total price");
    }
    return $price - $discount;
}

try {
    calculateDiscount(10.0, 15.0);
} catch (DomainException $e) {
    echo $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_length_exception_invalid_length() {
    compile_ok(
        r#"<?php
function validatePin(string $pin) {
    if (strlen($pin) !== 4) {
        throw new LengthException("PIN must be exactly 4 digits");
    }
}

try {
    validatePin("12");
} catch (LengthException $e) {
    echo $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_unexpected_value_exception_parsing() {
    compile_ok(
        r#"<?php
function parseFormat(string $format) {
    if ($format !== "json" && $format !== "xml") {
        throw new UnexpectedValueException("Expected json or xml, got $format");
    }
}

try {
    parseFormat("yaml");
} catch (UnexpectedValueException $e) {
    echo $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_range_exception_value_domain_error() {
    compile_ok(
        r#"<?php
function setPercentage(float $pct) {
    if ($pct < 0.0 || $pct > 1.0) {
        throw new RangeException("Percentage must be between 0.0 and 1.0");
    }
}

try {
    setPercentage(1.5);
} catch (RangeException $e) {
    echo $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_exception_hierarchy_instanceof_check() {
    compile_ok(
        r#"<?php
$e = new OutOfBoundsException("Index out of bounds");
echo ($e instanceof RuntimeException ? "RUNTIME_EX" : "NO");
echo ($e instanceof Exception ? " EXCEPTION" : " NO");
echo ($e instanceof Throwable ? " THROWABLE" : " NO");
"#,
    );
}

#[test]
fn test_php_type_error_builtin_exception() {
    compile_ok(
        r#"<?php
try {
    throw new TypeError("Type mismatch error");
} catch (Error $e) {
    echo "Caught Error: " . $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php_unhandled_match_error_php80() {
    compile_ok(
        r#"<?php
try {
    $x = 99;
    match ($x) {
        1 => "one",
        2 => "two",
    };
} catch (UnhandledMatchError $e) {
    echo "UnhandledMatchError: " . $e->getMessage();
}
"#,
    );
}
