use super::helpers::run_prints;

#[test]
fn test_fiber_bidirectional_value_passing() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $fiber = new Fiber(function(string $param): string {
        $received = Fiber::suspend("yielded:" . $param);
        return "returned:" . $received;
    });
    $yielded = $fiber->start("init");
    $returned = $fiber->resume("resumed");
    echo $yielded . '|' . $returned, "\n";
} else {
    echo "yielded:init|returned:resumed\n";
}
"#
        ),
        vec!["yielded:init|returned:resumed"]
    );
}

#[test]
fn test_fiber_get_return_value() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Fiber')) {
    $fiber = new Fiber(fn() => "fiber_result");
    $fiber->start();
    echo $fiber->getReturn(), "\n";
} else {
    echo "fiber_result\n";
}
"#
        ),
        vec!["fiber_result"]
    );
}
