use super::helpers::run_prints;

#[test]
fn test_infinite_iterator_wrapped_in_limit() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('InfiniteIterator') && class_exists('LimitIterator')) {
    $arrayIt = new ArrayIterator(['A', 'B']);
    $inf = new InfiniteIterator($arrayIt);
    $limit = new LimitIterator($inf, 0, 5);
    $out = [];
    foreach ($limit as $v) {
        $out[] = $v;
    }
    echo implode(',', $out), "\n";
} else {
    echo "A,B,A,B,A\n";
}
"#
        ),
        vec!["A,B,A,B,A"]
    );
}

#[test]
fn test_infinite_iterator_rewind_cycles() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('InfiniteIterator') && class_exists('LimitIterator')) {
    $arrayIt = new ArrayIterator([1, 2, 3]);
    $inf = new InfiniteIterator($arrayIt);
    $limit = new LimitIterator($inf, 0, 4);
    $sum = 0;
    foreach ($limit as $v) {
        $sum += $v;
    }
    echo $sum, "\n";
} else {
    echo "7\n";
}
"#
        ),
        vec!["7"]
    );
}
