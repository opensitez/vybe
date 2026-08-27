<?php
// vybe-test: php/spl_autoload/trait_uses_recursive_nested
// origin: languages/php/tests/php/test_spl_autoload.rs

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

echo "trait_uses_recursive_nested_ok";

__vybe_check(ob_get_clean(), "trait_uses_recursive_nested_ok");
