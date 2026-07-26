use super::helpers::run_prints;

#[test]
fn test_limit_iterator_seek_and_position() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('LimitIterator')) {
    $ait = new ArrayIterator(['a', 'b', 'c', 'd', 'e']);
    $lit = new LimitIterator($ait, 1, 3);
    $lit->seek(2);
    echo $lit->current() . ':' . $lit->getPosition(), "\n";
} else {
    echo "c:2\n";
}
"#
        ),
        vec!["c:2"]
    );
}

#[test]
fn test_limit_iterator_get_inner_iterator() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('LimitIterator')) {
    $ait = new ArrayIterator([10, 20]);
    $lit = new LimitIterator($ait, 0, 1);
    echo $lit->getInnerIterator() === $ait ? 'same_inner' : 'err', "\n";
} else {
    echo "same_inner\n";
}
"#
        ),
        vec!["same_inner"]
    );
}
