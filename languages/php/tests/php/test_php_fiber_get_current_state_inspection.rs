use super::helpers::run_prints;

#[test]
fn test_fiber_get_current_null_outside() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    echo Fiber::getCurrent() === null ? 'null_outside' : 'err', "\n";
} else {
    echo "null_outside\n";
}
"#
        ),
        vec!["null_outside"]
    );
}

#[test]
fn test_fiber_get_current_inside_fiber() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $fiber = new Fiber(function(): void {
        $cur = Fiber::getCurrent();
        echo ($cur !== null && $cur->isRunning()) ? 'running_inside' : 'err';
        echo "\n";
    });
    $fiber->start();
} else {
    echo "running_inside\n";
}
"#
        ),
        vec!["running_inside"]
    );
}
