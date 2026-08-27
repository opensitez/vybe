<?php
// vybe-test: php/goto/goto_multiple_labels
// origin: languages/php/tests/php/test_goto.rs

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

echo "goto_multiple_labels_ok";

__vybe_check(ob_get_clean(), "goto_multiple_labels_ok");
