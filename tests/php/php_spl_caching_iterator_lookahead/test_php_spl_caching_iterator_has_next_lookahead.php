<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_has_next_lookahead
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

$arr = new ArrayIterator(["first", "second", "third"]);
$it = new CachingIterator($arr);

$results = [];
foreach ($it as $val) {
    $results[] = $val . ":" . ($it->hasNext() ? "HAS_MORE" : "LAST");
}
echo implode(" | ", $results);

__vybe_check(ob_get_clean(), "first:HAS_MORE | second:HAS_MORE | third:LAST");
