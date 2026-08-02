<?php
// vybe-test: php/php_reflection_property_set_get_static/test_reflection_property_set_get_value_instance
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

class Container {
    public int $value = 10;
}
$c = new Container();
$rp = new ReflectionProperty(Container::class, 'value');
$rp->setValue($c, 50);
echo $rp->getValue($c), "\n";

__vybe_check(ob_get_clean(), "50");
