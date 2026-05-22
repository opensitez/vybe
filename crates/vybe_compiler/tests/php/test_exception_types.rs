use super::helpers::compile_ok;

// ── Exception construction ────────────────────────────────────

#[test] fn exception_with_message_and_code() {
    compile_ok(r#"<?php
$e = new Exception('not found', 404);
echo $e->getMessage();
echo $e->getCode();
"#);
}

#[test] fn exception_chained_previous() {
    compile_ok(r#"<?php
$cause = new RuntimeException('disk full');
$e = new Exception('write failed', 0, $cause);
echo $e->getPrevious()->getMessage();
"#);
}

// ── Custom hierarchy ──────────────────────────────────────────

#[test] fn custom_exception_hierarchy() {
    compile_ok(r#"<?php
class AppException extends RuntimeException {}
class NetworkException extends AppException {
    public function __construct(string $host) {
        parent::__construct('unreachable: ' . $host, 503);
    }
}
throw new NetworkException('example.com');
"#);
}

#[test] fn catch_parent_catches_child() {
    compile_ok(r#"<?php
class BaseException extends Exception {}
class ChildException extends BaseException {}
try {
    throw new ChildException('child');
} catch (BaseException $e) {
    echo $e->getMessage();
}
"#);
}

// ── Built-in SPL exception types ─────────────────────────────

#[test] fn runtime_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new RuntimeException('runtime issue');
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn invalid_argument_exception_builtin() {
    compile_ok(r#"<?php
function divide(int $a, int $b): float {
    if ($b === 0) throw new InvalidArgumentException('divisor cannot be zero');
    return $a / $b;
}
try { divide(10, 0); } catch (InvalidArgumentException $e) { echo $e->getMessage(); }
"#);
}

#[test] fn logic_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new LogicException('precondition violated');
} catch (LogicException $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn bad_method_call_exception_builtin() {
    compile_ok(r#"<?php
class Foo {
    public function bar(): void {
        throw new BadMethodCallException('bar not implemented');
    }
}
try { (new Foo())->bar(); } catch (BadMethodCallException $e) { echo $e->getMessage(); }
"#);
}

#[test] fn out_of_range_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new OutOfRangeException('index out of range');
} catch (OutOfRangeException $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn overflow_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new OverflowException('stack overflow');
} catch (OverflowException $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn underflow_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new UnderflowException('stack underflow');
} catch (UnderflowException $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn domain_exception_builtin() {
    compile_ok(r#"<?php
try {
    throw new DomainException('value outside domain');
} catch (DomainException $e) {
    echo $e->getMessage();
}
"#);
}

// ── PHP 7+ error types ────────────────────────────────────────

#[test] fn type_error_thrown() {
    compile_ok(r#"<?php
function strictAdd(int $a, int $b): int { return $a + $b; }
try {
    throw new TypeError('argument must be int');
} catch (TypeError $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn value_error_thrown() {
    compile_ok(r#"<?php
try {
    throw new ValueError('value out of acceptable range');
} catch (ValueError $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn arithmetic_error_builtin() {
    compile_ok(r#"<?php
try {
    throw new ArithmeticError('division undefined');
} catch (ArithmeticError $e) {
    echo $e->getMessage();
}
"#);
}

// ── Exception introspection methods ──────────────────────────

#[test] fn exception_get_message_code_line() {
    compile_ok(r#"<?php
try {
    throw new Exception('test message', 42);
} catch (Exception $e) {
    echo $e->getMessage();
    echo $e->getCode();
    echo $e->getLine();
}
"#);
}

// ── Re-throw and finally interactions ────────────────────────

#[test] fn rethrow_in_catch() {
    compile_ok(r#"<?php
function process(): void {
    try {
        throw new RuntimeException('original');
    } catch (RuntimeException $e) {
        throw new Exception('wrapped: ' . $e->getMessage(), 0, $e);
    }
}
try { process(); } catch (Exception $e) { echo $e->getMessage(); }
"#);
}

#[test] fn exception_in_finally() {
    compile_ok(r#"<?php
try {
    try {
        throw new Exception('first');
    } finally {
        throw new Exception('from finally');
    }
} catch (Exception $e) {
    echo $e->getMessage();
}
"#);
}

#[test] fn multiple_catch_different_types() {
    compile_ok(r#"<?php
function riskyOp(int $kind): void {
    if ($kind === 1) throw new InvalidArgumentException('bad arg');
    if ($kind === 2) throw new RuntimeException('runtime');
    throw new LogicException('logic');
}
foreach ([1, 2, 3] as $k) {
    try {
        riskyOp($k);
    } catch (InvalidArgumentException $e) {
        echo 'invalid';
    } catch (RuntimeException $e) {
        echo 'runtime';
    } catch (LogicException $e) {
        echo 'logic';
    }
}
"#);
}

// ── PHP 8 catch union type ────────────────────────────────────

#[test] fn catch_union_type() {
    compile_ok(r#"<?php
function risky(bool $flag): void {
    if ($flag) throw new TypeError('type');
    throw new ValueError('value');
}
foreach ([true, false] as $f) {
    try {
        risky($f);
    } catch (TypeError | ValueError $e) {
        echo $e->getMessage();
    }
}
"#);
}
