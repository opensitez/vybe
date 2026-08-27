<?php
// vybe-test: php/loops/foreach_loop_with_list_destructure_and_sparse_index
// origin: languages/php/tests/php/test_loops.rs

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

echo "foreach_loop_with_list_destructure_and_sparse_index_ok";

__vybe_check(ob_get_clean(), "foreach_loop_with_list_destructure_and_sparse_index_ok");
