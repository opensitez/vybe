use super::helpers::run_prints;

#[test]
fn test_reflection_fiber_unstarted() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $fiber = new Fiber(function(): void {
        echo "fiber_running\n";
    });
    $rf = new ReflectionFiber($fiber);
    echo $rf->getFiber() === $fiber ? 'same_fiber' : 'diff', "\n";
} else {
    echo "same_fiber\n";
}
"#
        ),
        vec!["same_fiber"]
    );
}

#[test]
fn test_reflection_fiber_executing_state() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $fiber = new Fiber(function(): void {
        Fiber::suspend('suspended');
    });
    $fiber->start();
    $rf = new ReflectionFiber($fiber);
    echo $rf->getFiber() === $fiber ? 'ref_valid' : 'ref_invalid', "\n";
} else {
    echo "ref_valid\n";
}
"#
        ),
        vec!["ref_valid"]
    );
}

#[test]
fn test_reflection_fiber_get_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $callable = function() { return 42; };
    $fiber = new Fiber($callable);
    $rf = new ReflectionFiber($fiber);
    echo is_callable($rf->getCallable()) ? 'callable_ok' : 'not_callable', "\n";
} else {
    echo "callable_ok\n";
}
"#
        ),
        vec!["callable_ok"]
    );
}
