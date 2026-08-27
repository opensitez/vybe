<?php
// vybe-test: php/php_spl_caching_iterator_full_cache/test_caching_iterator_has_next
// origin: languages/php/tests/php/test_php_spl_caching_iterator_full_cache.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_caching_iterator_has_next_ok";

__vybe_check(ob_get_clean(), "test_caching_iterator_has_next_ok");
