<?php
// vybe-test: php/loops/do_while_nested_continue_break_in_nested_switch
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

echo "do_while_nested_continue_break_in_nested_switch_ok";

__vybe_check(ob_get_clean(), "do_while_nested_continue_break_in_nested_switch_ok");
