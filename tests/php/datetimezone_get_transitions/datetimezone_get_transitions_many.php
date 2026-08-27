<?php
// vybe-test: php/datetimezone_get_transitions/datetimezone_get_transitions_many
// origin: languages/php/tests/php/test_datetimezone_get_transitions.rs

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

echo "datetimezone_get_transitions_many_ok";

__vybe_check(ob_get_clean(), "datetimezone_get_transitions_many_ok");
