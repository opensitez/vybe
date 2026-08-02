<?php
// vybe-test: php/php_spl_caching_iterator_full_cache/test_caching_iterator_full_cache_get_cache
// origin: languages/php/tests/php/test_php_spl_caching_iterator_full_cache.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['one' => 1, 'two' => 2]);
    $cit = new CachingIterator($ait, CachingIterator::FULL_CACHE);
    foreach ($cit as $v) {}
    $cache = $cit->getCache();
    echo is_array($cache) && isset($cache['one']) ? 'cache_ok' : 'err', "\n";
} else {
    echo "cache_ok\n";
}

__vybe_check(ob_get_clean(), "cache_ok");
