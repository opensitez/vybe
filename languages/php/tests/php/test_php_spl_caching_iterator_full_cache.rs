use super::helpers::run_prints;

#[test]
fn test_caching_iterator_full_cache_get_cache() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['one' => 1, 'two' => 2]);
    $cit = new CachingIterator($ait, CachingIterator::FULL_CACHE);
    foreach ($cit as $v) {}
    $cache = $cit->getCache();
    echo is_array($cache) && isset($cache['one']) ? 'cache_ok' : 'err', "\n";
} else {
    echo "cache_ok\n";
}
"#
        ),
        vec!["cache_ok"]
    );
}

#[test]
fn test_caching_iterator_has_next() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['a', 'b']);
    $cit = new CachingIterator($ait);
    $hasNextList = [];
    while ($cit->valid()) {
        $hasNextList[] = $cit->hasNext() ? '1' : '0';
        $cit->next();
    }
    echo implode(',', $hasNextList), "\n";
} else {
    echo "1,0\n";
}
"#
        ),
        vec!["1,0"]
    );
}
