<?php
// vybe-test: php/php_reflection_property_set_get_static/test_reflection_property_set_get_value_static
// origin: languages/php/tests/php/test_php_reflection_property_set_get_static.rs

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

class GlobalState {
    public static string $env = 'dev';
}
$rp = new ReflectionProperty(GlobalState::class, 'env');
$rp->setValue(null, 'prod');
echo $rp->getValue(), "\n";

__vybe_check(ob_get_clean(), "prod");
