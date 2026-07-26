use super::helpers::run_prints;

#[test]
fn test_spl_stack_top_and_bottom() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplStack')) {
    $stack = new SplStack();
    $stack->push('first');
    $stack->push('second');
    echo $stack->top() . ':' . $stack->bottom(), "\n";
} else {
    echo "second:first\n";
}
"#
        ),
        vec!["second:first"]
    );
}

#[test]
fn test_spl_stack_pop_order() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplStack')) {
    $stack = new SplStack();
    $stack->push(10);
    $stack->push(20);
    echo $stack->pop() . ',' . $stack->pop(), "\n";
} else {
    echo "20,10\n";
}
"#
        ),
        vec!["20,10"]
    );
}
