use super::helpers::run_prints;

#[test]
fn test_caching_iterator_tostring_use_current() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['hello', 'world']);
    $cit = new CachingIterator($ait, CachingIterator::TOSTRING_USE_CURRENT);
    echo (string)$cit, "\n";
} else {
    echo "hello\n";
}
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn test_caching_iterator_tostring_use_key() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['first' => 10, 'second' => 20]);
    $cit = new CachingIterator($ait, CachingIterator::TOSTRING_USE_KEY);
    echo (string)$cit, "\n";
} else {
    echo "first\n";
}
"#
        ),
        vec!["first"]
    );
}
