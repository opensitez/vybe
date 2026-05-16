use super::helpers::compile_ok;

// ── set_error_handler ────────────────────────────────────────

#[test] fn set_error_handler_basic() {
    compile_ok(r#"<?php
$errors = [];
set_error_handler(function(int $errno, string $errstr) use (&$errors): bool {
    $errors[] = "$errno: $errstr";
    return true; // suppress default handler
});
trigger_error("test warning", E_USER_WARNING);
restore_error_handler();
echo count($errors) > 0 ? 'caught' : 'missed';
"#);
}

#[test] fn set_error_handler_with_level() {
    compile_ok(r#"<?php
$caught = 0;
set_error_handler(function() use (&$caught): bool { $caught++; return true; }, E_USER_NOTICE);
trigger_error("a notice", E_USER_NOTICE);
trigger_error("a warning", E_USER_WARNING); // not caught by this handler
restore_error_handler();
echo $caught;
"#);
}

#[test] fn set_error_handler_full_signature() {
    compile_ok(r#"<?php
$last = [];
set_error_handler(function(int $errno, string $errstr, string $errfile, int $errline) use (&$last): bool {
    $last = ['no' => $errno, 'str' => $errstr, 'line' => $errline];
    return true;
});
trigger_error("custom error", E_USER_ERROR);
restore_error_handler();
echo $last['no'] === E_USER_ERROR ? 'correct errno' : 'wrong errno';
echo is_string($last['str']) ? ':has message' : ':no message';
"#);
}

#[test] fn set_error_handler_null_restores() {
    compile_ok(r#"<?php
set_error_handler(fn() => true);
$prev = set_error_handler(null); // restores default
echo 'restored';
"#);
}

// ── restore_error_handler ────────────────────────────────────

#[test] fn restore_error_handler_chain() {
    compile_ok(r#"<?php
$log = [];
set_error_handler(function(int $no, string $str) use (&$log): bool {
    $log[] = "H1:$str"; return true;
});
set_error_handler(function(int $no, string $str) use (&$log): bool {
    $log[] = "H2:$str"; return true;
});
trigger_error("msg", E_USER_NOTICE);
restore_error_handler(); // back to H1
trigger_error("msg2", E_USER_NOTICE);
restore_error_handler(); // back to default
echo count($log) . ':' . implode(',', $log);
"#);
}

// ── set_exception_handler ────────────────────────────────────

#[test] fn set_exception_handler_basic() {
    compile_ok(r#"<?php
$caught = null;
set_exception_handler(function(\Throwable $e) use (&$caught): void {
    $caught = $e->getMessage();
});
// Note: set_exception_handler catches uncaught exceptions at shutdown
// We test it's settable and callable
$prev = set_exception_handler(null); // restore
echo 'handler set';
"#);
}

// ── error_reporting ──────────────────────────────────────────

#[test] fn error_reporting_get() {
    compile_ok(r#"<?php
$current = error_reporting();
echo is_int($current) ? 'is int' : 'not int';
"#);
}

#[test] fn error_reporting_set() {
    compile_ok(r#"<?php
$old = error_reporting(E_ALL);
echo $old >= 0 ? 'got old value' : 'fail';
error_reporting($old); // restore
"#);
}

#[test] fn error_reporting_constants() {
    compile_ok(r#"<?php
echo E_ERROR       > 0 ? 'E_ERROR ok'   : 'fail';
echo E_WARNING     > 0 ? ':E_WARNING ok'   : ':fail';
echo E_NOTICE      > 0 ? ':E_NOTICE ok'    : ':fail';
echo E_DEPRECATED  > 0 ? ':E_DEPRECATED ok': ':fail';
echo E_USER_ERROR  > 0 ? ':E_USER_ERROR ok': ':fail';
echo E_ALL         > 0 ? ':E_ALL ok'       : ':fail';
"#);
}

#[test] fn error_reporting_bitmask() {
    compile_ok(r#"<?php
// Combine error levels with bitwise OR
$level = E_ERROR | E_WARNING | E_NOTICE;
$old = error_reporting($level);
echo error_reporting() === $level ? 'set correctly' : 'wrong';
error_reporting($old);
"#);
}

#[test] fn error_reporting_zero() {
    compile_ok(r#"<?php
$old = error_reporting(0);
echo error_reporting() === 0 ? 'suppressed' : 'not suppressed';
error_reporting($old);
"#);
}

// ── trigger_error ────────────────────────────────────────────

#[test] fn trigger_error_user_warning() {
    compile_ok(r#"<?php
$caught = false;
set_error_handler(function() use (&$caught): bool { $caught = true; return true; });
trigger_error("test", E_USER_WARNING);
restore_error_handler();
echo $caught ? 'triggered' : 'not triggered';
"#);
}

#[test] fn trigger_error_user_notice() {
    compile_ok(r#"<?php
$msg = '';
set_error_handler(function(int $no, string $str) use (&$msg): bool { $msg = $str; return true; });
trigger_error("hello notice", E_USER_NOTICE);
restore_error_handler();
echo $msg;
"#);
}

#[test] fn trigger_error_user_error() {
    compile_ok(r#"<?php
$caught = false;
set_error_handler(function() use (&$caught): bool { $caught = true; return true; });
trigger_error("fatal-like error", E_USER_ERROR);
restore_error_handler();
echo $caught ? 'caught' : 'missed';
"#);
}

#[test] fn trigger_error_user_deprecated() {
    compile_ok(r#"<?php
$caught = false;
set_error_handler(function(int $no) use (&$caught): bool {
    if ($no === E_USER_DEPRECATED) $caught = true;
    return true;
});
trigger_error("use newFunc() instead", E_USER_DEPRECATED);
restore_error_handler();
echo $caught ? 'deprecated caught' : 'missed';
"#);
}

// ── error_get_last ───────────────────────────────────────────

#[test] fn error_get_last_basic() {
    compile_ok(r#"<?php
set_error_handler(fn() => true); // suppress
@trigger_error("test error", E_USER_WARNING);
restore_error_handler();
$err = error_get_last();
echo $err !== null ? 'has error' : 'no error';
"#);
}

#[test] fn error_get_last_structure() {
    compile_ok(r#"<?php
set_error_handler(fn() => true);
trigger_error("structured error", E_USER_NOTICE);
restore_error_handler();
$err = error_get_last();
if ($err !== null) {
    echo isset($err['type'])    ? 'has type' : 'no type';
    echo isset($err['message']) ? ':has msg'  : ':no msg';
    echo isset($err['file'])    ? ':has file' : ':no file';
    echo isset($err['line'])    ? ':has line' : ':no line';
}
"#);
}

#[test] fn error_clear_last() {
    compile_ok(r#"<?php
set_error_handler(fn() => true);
trigger_error("something", E_USER_NOTICE);
restore_error_handler();
error_clear_last();
$err = error_get_last();
echo $err === null ? 'cleared' : 'still set';
"#);
}

// ── @ operator (error suppression) ───────────────────────────

#[test] fn at_operator_suppress() {
    compile_ok(r#"<?php
// @ suppresses errors from the expression
$result = @file_get_contents('/nonexistent/file/path');
echo $result === false ? 'failed silently' : 'unexpected success';
"#);
}

#[test] fn at_operator_with_handler() {
    compile_ok(r#"<?php
$triggered = false;
set_error_handler(function() use (&$triggered): bool { $triggered = true; return true; });
$r = @trigger_error("suppressed?", E_USER_NOTICE);
restore_error_handler();
// @ suppresses at the engine level — handler may or may not be called
echo is_bool($r) || $r === null ? 'ran' : 'fail';
"#);
}

// ── Throwable / Error class hierarchy ────────────────────────

#[test] fn error_vs_exception_hierarchy() {
    compile_ok(r#"<?php
try { throw new \Error("base error"); }
catch (\Throwable $t) { echo 'caught Throwable: ' . $t->getMessage(); }
"#);
}

#[test] fn type_error_catch() {
    compile_ok(r#"<?php
declare(strict_types=1);
function mustBeInt(int $n): int { return $n; }
try {
    $r = mustBeInt(3); // valid
    echo "ok: $r";
} catch (\TypeError $e) {
    echo 'type error: ' . $e->getMessage();
}
"#);
}

#[test] fn arithmetic_error() {
    compile_ok(r#"<?php
try { $r = intdiv(1, 0); }
catch (\DivisionByZeroError $e) { echo 'div by zero'; }
"#);
}

#[test] fn parse_error_via_eval() {
    compile_ok(r#"<?php
try {
    eval('$x = ;'); // parse error
} catch (\ParseError $e) {
    echo 'parse error caught';
}
"#);
}

// ── Exception chaining ────────────────────────────────────────

#[test] fn exception_previous() {
    compile_ok(r#"<?php
try {
    try {
        throw new \RuntimeException("original");
    } catch (\RuntimeException $e) {
        throw new \LogicException("wrapped", 0, $e);
    }
} catch (\LogicException $e) {
    echo $e->getMessage();
    echo ':' . $e->getPrevious()->getMessage();
}
"#);
}

#[test] fn exception_chain_deep() {
    compile_ok(r#"<?php
function buildChain(int $depth, ?\Throwable $prev = null): \Throwable {
    if ($depth === 0) return new \RuntimeException("root", 0, $prev);
    return buildChain($depth - 1, new \RuntimeException("level $depth", 0, $prev));
}
$e = buildChain(3);
$count = 0;
while ($e !== null) { $count++; $e = $e->getPrevious(); }
echo $count;
"#);
}

// ── Custom error handler class ────────────────────────────────

#[test] fn error_handler_class() {
    compile_ok(r#"<?php
class ErrorCollector {
    private array $errors = [];
    public function handle(int $errno, string $errstr): bool {
        $this->errors[] = ['no' => $errno, 'msg' => $errstr];
        return true;
    }
    public function getErrors(): array { return $this->errors; }
    public function count(): int { return count($this->errors); }
}
$collector = new ErrorCollector();
set_error_handler([$collector, 'handle']);
trigger_error("error one", E_USER_NOTICE);
trigger_error("error two", E_USER_WARNING);
restore_error_handler();
echo $collector->count();
echo ':' . $collector->getErrors()[0]['msg'];
"#);
}
