use super::helpers::compile_ok;

// ── Fiber creation ──────────────────────────────────────────
#[test]
fn fiber_new() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function() {
    echo "Hello from fiber";
});
"#,
    );
}

#[test]
fn fiber_start() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function() {
    echo "Running";
    return 42;
});
$result = $fiber->start();
"#,
    );
}

#[test]
fn fiber_suspend_resume() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function() {
    $value = Fiber::suspend('first');
    echo "Resumed with: " . $value;
    Fiber::suspend('second');
});
$v1 = $fiber->start();
echo $v1;
$v2 = $fiber->resume('hello');
echo $v2;
"#,
    );
}

#[test]
fn fiber_with_args() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function($x, $y) {
    Fiber::suspend($x + $y);
    return $x * $y;
});
$sum = $fiber->start(3, 4);
echo $sum;
"#,
    );
}

#[test]
fn fiber_state_check() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function() {
    Fiber::suspend();
});
$fiber->start();
$suspended = $fiber->isSuspended();
"#,
    );
}

#[test]
fn fiber_get_return() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function() {
    return 'done';
});
$fiber->start();
$result = $fiber->getReturn();
"#,
    );
}

// ── Generators (yield) ──────────────────────────────────────
#[test]
fn yield_basic() {
    compile_ok(
        r#"<?php
function gen() {
    yield 1;
    yield 2;
    yield 3;
}
"#,
    );
}

#[test]
fn yield_from_loop() {
    compile_ok(
        r#"<?php
function range_gen($start, $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
"#,
    );
}

#[test]
fn yield_no_value() {
    compile_ok(
        r#"<?php
function signals() {
    yield;
    yield;
}
"#,
    );
}

#[test]
fn yield_key_value() {
    compile_ok(
        r#"<?php
function pairs() {
    yield 'a';
    yield 'b';
    yield 'c';
}
"#,
    );
}

#[test]
fn yield_in_closure() {
    compile_ok(
        r#"<?php
$gen = function() {
    yield 1;
    yield 2;
};
"#,
    );
}

// ── Fiber + closure capture ─────────────────────────────────
#[test]
fn fiber_captures_variable() {
    compile_ok(
        r#"<?php
$message = "Hello";
$fiber = new Fiber(function() use ($message) {
    Fiber::suspend($message . " World");
});
$result = $fiber->start();
echo $result;
"#,
    );
}

// ── Multiple fibers ─────────────────────────────────────────
#[test]
fn multiple_fibers() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

// ── Generator as iterator pattern ───────────────────────────
#[test]
fn generator_fibonacci() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn yield_from_expr() {
    compile_ok(
        r#"<?php
function inner() {
    yield 1;
    yield 2;
}
function outer() {
    yield from inner();
    yield 3;
}
"#,
    );
}

// ── Runtime fiber suspend/resume (`php_cases!`) ─────────────────

crate::php_cases! {
    fiber_start_returns_first_suspend_value => {
        r#"<?php
$f = new Fiber(function (): void {
    Fiber::suspend('pause');
    echo 'after';
});
echo $f->start();
"#,
        ["pause"]
    };

    fiber_resume_passes_value_back_into_fiber => {
        r#"<?php
$f = new Fiber(function (): void {
    $v = Fiber::suspend('need');
    echo $v;
});
$f->start();
$f->resume('data');
"#,
        ["data"]
    };

    fiber_return_value_available_after_completion => {
        r#"<?php
$f = new Fiber(function (): int {
    Fiber::suspend('mid');
    return 99;
});
$f->start();
echo $f->resume();
"#,
        ["99"]
    };

    fiber_is_started_after_start => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend(1); });
$f->start();
echo $f->isStarted() ? 'started' : 'new';
"#,
        ["started"]
    };

    fiber_is_suspended_while_waiting_for_resume => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend('x'); });
$f->start();
echo $f->isSuspended() ? 'wait' : 'run';
"#,
        ["wait"]
    };

    fiber_is_terminated_after_return => {
        r#"<?php
$f = new Fiber(function (): int { return 1; });
$f->start();
echo $f->isTerminated() ? 'done' : 'live';
"#,
        ["done"]
    };

    fiber_get_return_throws_while_suspended => {
        r#"<?php
$f = new Fiber(function (): int {
    Fiber::suspend('hold');
    return 5;
});
$f->start();
try { $f->getReturn(); echo 'got'; } catch (FiberError) { echo 'blocked'; }
"#,
        ["blocked"]
    };

    fiber_throws_exception_propagates_to_caller => {
        r#"<?php
$f = new Fiber(function (): void { throw new RuntimeException('boom'); });
try { $f->start(); echo 'ok'; } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["boom"]
    };

    fiber_suspend_outside_fiber_throws => {
        r#"<?php
try { Fiber::suspend('nope'); echo 'ok'; } catch (FiberError) { echo 'err'; }
"#,
        ["err"]
    };

    fiber_resume_before_start_throws => {
        r#"<?php
$f = new Fiber(function (): void {});
try { $f->resume(); echo 'ok'; } catch (FiberError) { echo 'early'; }
"#,
        ["early"]
    };

    fiber_start_accepts_arguments_passed_to_closure => {
        r#"<?php
$f = new Fiber(function (int $a, int $b): void {
    echo $a + $b;
});
$f->start(2, 5);
"#,
        ["7"]
    };

    fiber_nested_suspend_resume_sequence => {
        r#"<?php
$f = new Fiber(function (): void {
    echo Fiber::suspend('a');
    echo Fiber::suspend('b');
    echo 'end';
});
$f->start();
$f->resume('1');
$f->resume('2');
"#,
        ["12end"]
    };
}
