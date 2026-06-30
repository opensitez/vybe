//! Fiber invalid-state and out-of-context errors — not covered by `test_fiber_lifecycle.rs`.

crate::php_cases! {
    fiber_suspend_outside_fiber_throws_error => {
        r#"<?php
try { Fiber::suspend(); echo 'ok'; }
catch (Error $e) { echo 'no-fiber'; }
"#,
        ["no-fiber"]
    };

    fiber_get_current_outside_is_null => {
        r#"<?php
echo Fiber::getCurrent() === null ? 'null' : 'obj';
"#,
        ["null"]
    };

    fiber_resume_before_start_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend('x'); });
try { $f->resume(); echo 'ok'; }
catch (FiberError $e) { echo 'early-resume'; }
"#,
        ["early-resume"]
    };

    fiber_start_twice_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend(); });
$f->start();
try { $f->start(); echo 'ok'; }
catch (FiberError $e) { echo 'twice'; }
"#,
        ["twice"]
    };

    fiber_throw_before_start_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend(); });
try { $f->throw(new Exception('x')); echo 'ok'; }
catch (FiberError $e) { echo 'early-throw'; }
"#,
        ["early-throw"]
    };

    fiber_uncaught_exception_from_start_bubbles_to_caller => {
        r#"<?php
$f = new Fiber(function (): void { throw new RuntimeException('start-boom'); });
try { $f->start(); echo 'ok'; }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["start-boom"]
    };

    fiber_uncaught_exception_on_resume_bubbles_to_caller => {
        r#"<?php
$f = new Fiber(function (): void {
    Fiber::suspend('pause');
    throw new LogicException('resume-boom');
});
$f->start();
try { $f->resume(); echo 'ok'; }
catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["resume-boom"]
    };

    fiber_get_return_before_completion_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend(1); });
$f->start();
try { $f->getReturn(); echo 'ok'; }
catch (FiberError $e) { echo 'no-return'; }
"#,
        ["no-return"]
    };

    fiber_resume_after_normal_completion_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): int { return 7; });
$f->start();
try { $f->resume(); echo 'ok'; }
catch (FiberError $e) { echo 'done-resume'; }
"#,
        ["done-resume"]
    };

    fiber_is_terminated_after_uncaught_inner_exception => {
        r#"<?php
$f = new Fiber(function (): void { throw new Exception('die'); });
try { $f->start(); } catch (Exception $e) { /* handled */ }
echo $f->isTerminated() ? 'dead' : 'alive';
"#,
        ["dead"]
    };

    fiber_throw_after_termination_throws_fiber_error => {
        r#"<?php
$f = new Fiber(function (): void {});
$f->start();
try { $f->throw(new Exception('late')); echo 'ok'; }
catch (FiberError $e) { echo 'late-throw'; }
"#,
        ["late-throw"]
    };

    fiber_suspend_value_visible_on_resume => {
        r#"<?php
$f = new Fiber(function (): void {
    $n = Fiber::suspend('payload');
    echo $n;
});
$f->start();
$f->resume(42);
"#,
        ["42"]
    };

    fiber_nested_get_current_identifies_running_fiber => {
        r#"<?php
$outer = new Fiber(function (): void {
    $inner = new Fiber(function () use (&$inner): void {
        $cur = Fiber::getCurrent();
        echo ($cur === $inner) ? 'inner' : 'other';
    });
    $inner->start();
});
$outer->start();
"#,
        ["inner"]
    };

    fiber_callback_must_be_callable_type_error => {
        r#"<?php
try { new Fiber(123); echo 'ok'; }
catch (TypeError $e) { echo 'not-callable'; }
"#,
        ["not-callable"]
    };

    fiber_is_suspended_false_before_start => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend(); });
echo $f->isSuspended() ? 'yes' : 'no';
"#,
        ["no"]
    };

    fiber_is_running_false_from_outside_while_suspended => {
        r#"<?php
$f = new Fiber(function (): void { Fiber::suspend('hold'); });
$f->start();
echo $f->isRunning() ? 'run' : 'idle';
"#,
        ["idle"]
    };
}
