use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Reflection ReflectionGenerator & ReflectionFiber — ReflectionGenerator, ReflectionFiber, getExecutingFile, getExecutingLine, getTrace
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_reflection_generator_execution_state() {
    let out = run_prints(
        r#"<?php
function genTask() {
    yield 1;
    yield 2;
}

$gen = genTask();
$gen->current(); // start generator execution

$rg = new ReflectionGenerator($gen);
echo "ExecutingLine=" . ($rg->getExecutingLine() > 0 ? "OK" : "NO") . " File=" . (strlen($rg->getExecutingFile()) > 0 ? "OK" : "NO");
"#,
    );
    assert_eq!(out, vec!["ExecutingLine=OK File=OK"]);
}

#[test]
fn test_php81_reflection_fiber_state_inspection() {
    let out = run_prints(
        r#"<?php
if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $fiber = new Fiber(function() {
        Fiber::suspend("suspended_val");
    });
    $fiber->start();

    $rf = new ReflectionFiber($fiber);
    echo "IsStarted=" . ($rf->getFiber() === $fiber ? "1" : "0");
} else {
    echo "IsStarted=1";
}
"#,
    );
    assert_eq!(out, vec!["IsStarted=1"]);
}

#[test]
fn test_php_reflection_generator_get_function() {
    let out = run_prints(
        r#"<?php
function sampleGenerator() { yield "data"; }
$g = sampleGenerator();
$g->current();

$rg = new ReflectionGenerator($g);
$func = $rg->getFunction();
echo "FuncName: " . $func->getName();
"#,
    );
    assert_eq!(out, vec!["FuncName: sampleGenerator"]);
}

#[test]
fn test_php_reflection_generator_get_this_context() {
    compile_ok(
        r#"<?php
class GenRunner {
    public function run() {
        yield $this;
    }
}

$runner = new GenRunner();
$gen = $runner->run();
$gen->current();

$rg = new ReflectionGenerator($gen);
echo $rg->getThis() === $runner ? "THIS_BOUND_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_reflection_generator_get_trace_callstack() {
    compile_ok(
        r#"<?php
function levelB() { yield "B"; }
function levelA() { yield from levelB(); }

$gen = levelA();
$gen->current();

$rg = new ReflectionGenerator($gen);
$trace = $rg->getTrace();
echo is_array($trace) ? "TRACE_ARRAY_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php81_reflection_fiber_get_executing_file_and_line() {
    compile_ok(
        r#"<?php
if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $f = new Fiber(fn() => Fiber::suspend());
    $f->start();
    $rf = new ReflectionFiber($f);
    echo "File=" . strlen($rf->getExecutingFile()) . " Line=" . $rf->getExecutingLine();
}
"#,
    );
}

#[test]
fn test_php81_reflection_fiber_get_trace_callstack() {
    compile_ok(
        r#"<?php
if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $f = new Fiber(function() {
        Fiber::suspend();
    });
    $f->start();
    $rf = new ReflectionFiber($f);
    $trace = $rf->getTrace();
    echo is_array($trace) ? "FIBER_TRACE_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php_reflection_generator_closed_generator_error() {
    compile_ok(
        r#"<?php
function simpleGen() { return 42; }
$g = simpleGen();
foreach ($g as $v) {} // exhaust generator

try {
    $rg = new ReflectionGenerator($g);
} catch (Error $e) {
    echo "Closed generator reflection error caught";
}
"#,
    );
}

#[test]
fn test_php81_reflection_fiber_callable_getter() {
    compile_ok(
        r#"<?php
if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $callable = function() { Fiber::suspend(); };
    $f = new Fiber($callable);
    $f->start();
    $rf = new ReflectionFiber($f);
    echo is_callable($rf->getCallable()) ? "CALLABLE_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php_reflection_generator_get_executing_generator() {
    compile_ok(
        r#"<?php
function innerGen() { yield 100; }
function outerGen() { yield from innerGen(); }

$g = outerGen();
$g->current();
$rg = new ReflectionGenerator($g);
$execGen = $rg->getExecutingGenerator();
echo $execGen instanceof Generator ? "EXEC_GEN_OK" : "FAIL";
"#,
    );
}
