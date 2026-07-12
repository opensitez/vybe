use super::helpers::run_prints;

// ── Fiber basic start/resume ──────────────────────────────────

#[test]
fn fiber_start_returns_first_yielded_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    $val = Fiber::suspend('first');
    echo "got: $val";
});
$result = $fiber->start();
echo $result;
"#
        ),
        vec!["first"]
    );
}

#[test]
fn fiber_resume_sends_value_and_gets_next_suspend() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    $a = Fiber::suspend(1);
    $b = Fiber::suspend(2);
    echo "$a,$b";
});
$fiber->start();
$fiber->resume('x');
$fiber->resume('y');
"#
        ),
        vec!["x,y"]
    );
}

#[test]
fn fiber_multiple_suspend_resume_cycles() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    $sum = 0;
    while (true) {
        $n = Fiber::suspend($sum);
        if ($n === null) break;
        $sum += $n;
    }
});
$fiber->start();
$fiber->resume(10);
$fiber->resume(20);
echo $fiber->resume(5);
"#
        ),
        vec!["35"]
    );
}

// ── Fiber status: isStarted ───────────────────────────────────

#[test]
fn fiber_is_not_started_before_start() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
echo $fiber->isStarted() ? 'started' : 'not started';
"#
        ),
        vec!["not started"]
    );
}

#[test]
fn fiber_is_started_after_start() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
$fiber->start();
echo $fiber->isStarted() ? 'started' : 'not started';
"#
        ),
        vec!["started"]
    );
}

// ── Fiber status: isSuspended ─────────────────────────────────

#[test]
fn fiber_is_suspended_after_first_suspend() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
$fiber->start();
echo $fiber->isSuspended() ? 'suspended' : 'not suspended';
"#
        ),
        vec!["suspended"]
    );
}

#[test]
fn fiber_is_not_suspended_when_running() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    echo Fiber::getCurrent()->isSuspended() ? 'yes' : 'no';
    Fiber::suspend();
});
$fiber->start();
"#
        ),
        vec!["no"]
    );
}

// ── Fiber status: isRunning ───────────────────────────────────

#[test]
fn fiber_is_running_from_within() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    echo Fiber::getCurrent()->isRunning() ? 'running' : 'not';
    Fiber::suspend();
});
$fiber->start();
"#
        ),
        vec!["running"]
    );
}

// ── Fiber status: isTerminated ────────────────────────────────

#[test]
fn fiber_is_terminated_after_completion() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { /* no suspend */ });
$fiber->start();
echo $fiber->isTerminated() ? 'terminated' : 'alive';
"#
        ),
        vec!["terminated"]
    );
}

#[test]
fn fiber_is_not_terminated_while_suspended() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
$fiber->start();
echo $fiber->isTerminated() ? 'terminated' : 'alive';
"#
        ),
        vec!["alive"]
    );
}

// ── Fiber::getCurrent ─────────────────────────────────────────

#[test]
fn fiber_get_current_returns_null_outside_fiber() {
    assert_eq!(
        run_prints(
            r#"<?php
echo var_export(Fiber::getCurrent(), true);
"#
        ),
        vec!["NULL"]
    );
}

#[test]
fn fiber_get_current_returns_self_inside() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    echo (Fiber::getCurrent() !== null) ? 'has current' : 'no current';
    Fiber::suspend();
});
$fiber->start();
"#
        ),
        vec!["has current"]
    );
}

// ── Fiber return value ────────────────────────────────────────

#[test]
fn fiber_get_return_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): int {
    Fiber::suspend();
    return 42;
});
$fiber->start();
$fiber->resume();
echo $fiber->getReturn();
"#
        ),
        vec!["42"]
    );
}

#[test]
fn fiber_return_value_null_when_no_explicit_return() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
$fiber->start();
$fiber->resume();
echo var_export($fiber->getReturn(), true);
"#
        ),
        vec!["NULL"]
    );
}

// ── Fiber exception handling ──────────────────────────────────

#[test]
fn fiber_throw_sends_exception_into_fiber() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    try {
        Fiber::suspend();
    } catch (RuntimeException $e) {
        echo "caught: " . $e->getMessage();
    }
});
$fiber->start();
$fiber->throw(new RuntimeException("from outside"));
"#
        ),
        vec!["caught: from outside"]
    );
}

#[test]
fn fiber_uncaught_exception_propagates_to_caller() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend(); });
$fiber->start();
try {
    $fiber->throw(new LogicException("logic error"));
} catch (LogicException $e) {
    echo $e->getMessage();
}
"#
        ),
        vec!["logic error"]
    );
}

// ── Fiber cooperative multitasking pattern ────────────────────

#[test]
fn two_fibers_interleaved_execution() {
    assert_eq!(
        run_prints(
            r#"<?php
$log = [];
$a = new Fiber(function() use (&$log): void {
    $log[] = 'A1'; Fiber::suspend();
    $log[] = 'A2'; Fiber::suspend();
    $log[] = 'A3';
});
$b = new Fiber(function() use (&$log): void {
    $log[] = 'B1'; Fiber::suspend();
    $log[] = 'B2';
});
$a->start(); $b->start();
$a->resume(); $b->resume();
$a->resume();
echo implode(',', $log);
"#
        ),
        vec!["A1,B1,A2,B2,A3"]
    );
}

// ── Fiber as async-like coroutine ─────────────────────────────

#[test]
fn fiber_simulates_async_task_queue() {
    assert_eq!(
        run_prints(
            r#"<?php
$tasks = [];
$tasks[] = new Fiber(function(): void { echo "task1:start\n"; Fiber::suspend(); echo "task1:end\n"; });
$tasks[] = new Fiber(function(): void { echo "task2:start\n"; Fiber::suspend(); echo "task2:end\n"; });
foreach ($tasks as $t) $t->start();
foreach ($tasks as $t) if ($t->isSuspended()) $t->resume();
"#
        ),
        vec!["task1:start", "task2:start", "task1:end", "task2:end"]
    );
}

// ── Fiber value passing through suspend chain ─────────────────

#[test]
fn fiber_pass_values_both_directions() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): string {
    $x = Fiber::suspend("need input");
    $y = Fiber::suspend("got: $x");
    return "final: " . ($x + $y);
});
$prompt1 = $fiber->start();
echo $prompt1 . "\n";
$prompt2 = $fiber->resume(10);
echo $prompt2 . "\n";
$fiber->resume(5);
echo $fiber->getReturn();
"#
        ),
        vec!["need input", "got: 10", "final: 15"]
    );
}

// ── Fiber resuming after exception caught inside ──────────────

#[test]
fn fiber_continues_after_catching_thrown_exception() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    try { Fiber::suspend(); } catch (\Exception $e) {}
    Fiber::suspend('after catch');
});
$fiber->start();
$fiber->throw(new \Exception("x"));
echo $fiber->getReturn() ?? $fiber->resume();
"#
        ),
        vec!["after catch"]
    );
}

// ── Fiber in class context ────────────────────────────────────

#[test]
fn fiber_inside_class_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Worker {
    public function run(): Fiber {
        return new Fiber(function(): void {
            $result = Fiber::suspend('ready');
            echo "processed: $result";
        });
    }
}
$fiber = (new Worker())->run();
$fiber->start();
$fiber->resume('input');
"#
        ),
        vec!["processed: input"]
    );
}

// ── Fiber cannot be resumed once terminated ───────────────────

#[test]
fn fiber_throw_on_resume_after_termination() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {});
$fiber->start();
try {
    $fiber->resume();
} catch (FiberError $e) {
    echo "fiber error";
}
"#
        ),
        vec!["fiber error"]
    );
}
