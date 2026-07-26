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

    fiber_start_only_once_throws_on_second_start => {
        r#"<?php
$f = new Fiber(function (): void {
    Fiber::suspend('first');
});
$f->start();
try {
    $f->start();
    echo 'double';
} catch (FiberError) {
    echo 'blocked';
}
"#,
        ["blocked"]
    };

    fiber_resume_when_terminated_throws => {
        r#"<?php
$f = new Fiber(function (): int { return 4; });
$f->start();
try {
    $f->resume();
    echo 'ok';
} catch (FiberError $e) {
    echo 'stopped';
}
"#,
        ["stopped"]
    };

    fiber_throw_from_within_calls_caller_exception => {
        r#"<?php
$f = new Fiber(function (): void {
    throw new Exception('inside');
});
try {
    $f->start();
    echo 'unreached';
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["inside"]
    };

    fiber_suspend_can_return_null => {
        r#"<?php
$f = new Fiber(function (): void {
    $v = Fiber::suspend();
    echo $v === null ? 'null' : 'set';
});
$f->start();
$f->resume('payload');
"#,
        ["nullnull"]
    };

    fiber_nested_within_fiber => {
        r#"<?php
$inner = new Fiber(function (): string {
    return 'inner';
});
$outer = new Fiber(function () use ($inner): string {
    $inner->start();
    return $inner->getReturn() . '-outer';
});
$outer->start();
echo $outer->getReturn();
"#,
        ["inner-outer"]
    };

    generator_yield_key_preserved_in_foreach => {
        r#"<?php
function gen_keys() {
    yield 0 => 'zero';
    yield 2 => 'two';
}
$out = [];
foreach (gen_keys() as $k => $v) {
    $out[] = $k . ':' . $v;
}
echo implode('|', $out);
"#,
        ["0:zero|2:two"]
    };

    generator_send_after_suspend => {
        r#"<?php
function gen_values() {
    $first = yield 'a';
    echo $first;
    yield $first . 'b';
}
$g = gen_values();
$x = $g->current();
echo $x;
$x = $g->send('z');
echo $x;
"#,
        ["azb"]
    };

    generator_rewind_and_iter => {
        r#"<?php
function nums() { yield 1; yield 2; yield 3; }
$g = nums();
$vals = [];
$g->rewind();
while ($g->valid()) {
    $vals[] = $g->current();
    $g->next();
}
echo implode(',', $vals);
"#,
        ["1,2,3"]
    };

    fiber_double_resume_without_start_is_error => {
        r#"<?php
$f = new Fiber(function (): void {
    Fiber::suspend('first');
});
$f->start();
echo '|';
try {
    $f->resume();
    echo 'running';
} catch (FiberError $e) {
    echo 'err1';
}
$f->resume();
echo '|';
try {
    $f->resume();
    echo 'second';
} catch (FiberError $e) {
    echo 'stopped';
}
"#,
        ["first|running|stopped"]
    };

    fiber_get_return_after_exception_is_error => {
        r#"<?php
$f = new Fiber(function (): void {
    throw new RuntimeException('boom');
});
try {
    $f->start();
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
try {
    $f->getReturn();
    echo '|bad';
} catch (FiberError) {
    echo '|blocked';
}
"#,
        ["boom|blocked"]
    };

    fiber_resume_to_throw_exception_to_fiber => {
        r#"<?php
$f = new Fiber(function (): void {
    try {
        Fiber::suspend('enter');
    } catch (Throwable $e) {
        echo 'caught:' . $e->getMessage();
    }
});
echo $f->start();
try {
    $f->throw(new RuntimeException('from-caller'));
} catch (FiberError) {
    echo '|throw-failed';
}
"#,
        ["enter|caught:from-caller"]
    };

    fiber_suspend_then_terminate_flow => {
        r#"<?php
$f = new Fiber(function (): int {
    echo Fiber::suspend('a');
    Fiber::suspend('b');
    return 42;
});
echo $f->start();
echo '|';
echo $f->resume('x');
echo '|';
echo $f->resume('y');
"#,
        ["a|x|42"]
    };

    generator_send_after_start => {
        r#"<?php
function chain() {
    $v = yield 'start';
    yield $v . '-next';
}
$g = chain();
echo $g->current();
echo '|';
echo $g->send('sent');
"#,
        ["start|sent-next"]
    };
}
