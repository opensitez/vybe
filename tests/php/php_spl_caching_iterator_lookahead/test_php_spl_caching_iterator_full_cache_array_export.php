<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_full_cache_array_export
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs

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

$arr = new ArrayIterator(["x" => 10, "y" => 20]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);

foreach ($it as $val) {}

$cache = $it->getCache();
echo "Cache count=" . count($cache) . " Y=" . $cache["y"];

__vybe_check(ob_get_clean(), "Cache count=2 Y=20");
