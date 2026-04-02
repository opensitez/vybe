mod helpers;
use helpers::compile_ok;

// ── Fiber creation ──────────────────────────────────────────
#[test] fn fiber_new() { compile_ok(r#"<?php
$fiber = new Fiber(function() {
    echo "Hello from fiber";
});
"#); }

#[test] fn fiber_start() { compile_ok(r#"<?php
$fiber = new Fiber(function() {
    echo "Running";
    return 42;
});
$result = $fiber->start();
"#); }

#[test] fn fiber_suspend_resume() { compile_ok(r#"<?php
$fiber = new Fiber(function() {
    $value = Fiber::suspend('first');
    echo "Resumed with: " . $value;
    Fiber::suspend('second');
});
$v1 = $fiber->start();
echo $v1;
$v2 = $fiber->resume('hello');
echo $v2;
"#); }

#[test] fn fiber_with_args() { compile_ok(r#"<?php
$fiber = new Fiber(function($x, $y) {
    Fiber::suspend($x + $y);
    return $x * $y;
});
$sum = $fiber->start(3, 4);
echo $sum;
"#); }

#[test] fn fiber_state_check() { compile_ok(r#"<?php
$fiber = new Fiber(function() {
    Fiber::suspend();
});
$fiber->start();
$suspended = $fiber->isSuspended();
"#); }

#[test] fn fiber_get_return() { compile_ok(r#"<?php
$fiber = new Fiber(function() {
    return 'done';
});
$fiber->start();
$result = $fiber->getReturn();
"#); }

// ── Generators (yield) ──────────────────────────────────────
#[test] fn yield_basic() { compile_ok(r#"<?php
function gen() {
    yield 1;
    yield 2;
    yield 3;
}
"#); }

#[test] fn yield_from_loop() { compile_ok(r#"<?php
function range_gen($start, $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
"#); }

#[test] fn yield_no_value() { compile_ok(r#"<?php
function signals() {
    yield;
    yield;
}
"#); }

#[test] fn yield_key_value() { compile_ok(r#"<?php
function pairs() {
    yield 'a';
    yield 'b';
    yield 'c';
}
"#); }

#[test] fn yield_in_closure() { compile_ok(r#"<?php
$gen = function() {
    yield 1;
    yield 2;
};
"#); }

// ── Fiber + closure capture ─────────────────────────────────
#[test] fn fiber_captures_variable() { compile_ok(r#"<?php
$message = "Hello";
$fiber = new Fiber(function() use ($message) {
    Fiber::suspend($message . " World");
});
$result = $fiber->start();
echo $result;
"#); }

// ── Multiple fibers ─────────────────────────────────────────
#[test] fn multiple_fibers() { compile_ok(r#"<?php
$f1 = new Fiber(function() {
    Fiber::suspend('f1');
    return 'f1 done';
});
$f2 = new Fiber(function() {
    Fiber::suspend('f2');
    return 'f2 done';
});
echo $f1->start();
echo $f2->start();
$f1->resume();
$f2->resume();
"#); }

// ── Generator as iterator pattern ───────────────────────────
#[test] fn generator_fibonacci() { compile_ok(r#"<?php
function fibonacci() {
    $a = 0;
    $b = 1;
    while (true) {
        yield $a;
        $tmp = $a;
        $a = $b;
        $b = $tmp + $b;
    }
}
"#); }

#[test] fn yield_from_expr() { compile_ok(r#"<?php
function inner() {
    yield 1;
    yield 2;
}
function outer() {
    yield from inner();
    yield 3;
}
"#); }
