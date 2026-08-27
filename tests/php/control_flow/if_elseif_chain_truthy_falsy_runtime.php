<?php
// vybe-test: php/control_flow/if_elseif_chain_truthy_falsy_runtime
// origin: languages/php/tests/php/test_control_flow.rs

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

echo "if_elseif_chain_truthy_falsy_runtime_ok";

__vybe_check(ob_get_clean(), "if_elseif_chain_truthy_falsy_runtime_ok");
