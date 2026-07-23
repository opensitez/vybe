use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Fibers & Asynchronous Concurrency — Fiber, suspend, resume, start, isStarted, isTerminated, throw
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php81_fiber_basic_suspend_and_resume() {
    let out = run_prints(
        r#"<?php
$fiber = new Fiber(function(): void {
    $value = Fiber::suspend("fiber_yield_1");
    echo "received: $value";
});

$res1 = $fiber->start();
echo "start=$res1 | ";
$fiber->resume("resume_val");
"#,
    );
    assert_eq!(out, vec!["start=fiber_yield_1 | received: resume_val"]);
}

#[test]
fn test_php81_fiber_state_predicates() {
    let out = run_prints(
        r#"<?php
$fiber = new Fiber(function(): int {
    Fiber::suspend();
    return 42;
});

echo $fiber->isStarted() ? "0" : "1";
$fiber->start();
echo $fiber->isSuspended() ? "1" : "0";
$res = $fiber->resume();
echo $fiber->isTerminated() ? "1" : "0";
echo " res=$res";
"#,
    );
    assert_eq!(out, vec!["111 res=42"]);
}

#[test]
fn test_php81_fiber_exception_injection_throw() {
    let out = run_prints(
        r#"<?php
$fiber = new Fiber(function(): void {
    try {
        Fiber::suspend();
    } catch (RuntimeException $e) {
        echo "CAUGHT_IN_FIBER: " . $e->getMessage();
    }
});

$fiber->start();
$fiber->throw(new RuntimeException("Fiber Exception"));
"#,
    );
    assert_eq!(out, vec!["CAUGHT_IN_FIBER: Fiber Exception"]);
}

#[test]
fn test_php81_fiber_return_value_get_return() {
    let out = run_prints(
        r#"<?php
$fiber = new Fiber(fn() => "fiber_result");
$fiber->start();
echo $fiber->getReturn();
"#,
    );
    assert_eq!(out, vec!["fiber_result"]);
}

#[test]
fn test_php81_fiber_current_reference() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function(): void {
    $curr = Fiber::getCurrent();
    echo $curr !== null ? "HAS_CURRENT" : "NO_CURRENT";
});
$fiber->start();
"#,
    );
}

#[test]
fn test_php81_fiber_nested_fiber_execution() {
    compile_ok(
        r#"<?php
$f1 = new Fiber(function(): void {
    $f2 = new Fiber(function(): void {
        Fiber::suspend("inner_yield");
    });
    $val = $f2->start();
    echo "F1 received $val";
});
$f1->start();
"#,
    );
}

#[test]
fn test_php81_fiber_cannot_suspend_outside_fiber_error() {
    compile_ok(
        r#"<?php
try {
    Fiber::suspend();
} catch (FiberError $e) {
    echo "FiberError: " . $e->getMessage();
}
"#,
    );
}

#[test]
fn test_php81_fiber_cannot_resume_terminated_error() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(fn() => 123);
$fiber->start();
try {
    $fiber->resume();
} catch (FiberError $e) {
    echo "Cannot resume terminated fiber";
}
"#,
    );
}

#[test]
fn test_php81_fiber_event_loop_cooperative_scheduler() {
    compile_ok(
        r#"<?php
class SimpleLoop {
    private array $fibers = [];
    public function enqueue(Fiber $f): void { $this->fibers[] = $f; }
    public function run(): void {
        while (!empty($this->fibers)) {
            $f = array_shift($this->fibers);
            if (!$f->isStarted()) { $f->start(); }
            elseif ($f->isSuspended()) { $f->resume(); }
            if ($f->isSuspended()) { $this->fibers[] = $f; }
        }
    }
}

$loop = new SimpleLoop();
$loop->enqueue(new Fiber(function() { echo "Task 1\n"; Fiber::suspend(); echo "Task 1 Done\n"; }));
$loop->enqueue(new Fiber(function() { echo "Task 2\n"; Fiber::suspend(); echo "Task 2 Done\n"; }));
$loop->run();
"#,
    );
}

#[test]
fn test_php81_fiber_argument_passing_on_start() {
    compile_ok(
        r#"<?php
$fiber = new Fiber(function(string $name, int $id): string {
    return "$name#$id";
});
$res = $fiber->start("Worker", 99);
echo $res;
"#,
    );
}
