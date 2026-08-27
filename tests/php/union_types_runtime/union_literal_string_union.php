<?php
// vybe-test: php/union_types_runtime/union_literal_string_union
// origin: languages/php/tests/php/test_union_types_runtime.rs

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

echo "union_literal_string_union_ok";

__vybe_check(ob_get_clean(), "union_literal_string_union_ok");
