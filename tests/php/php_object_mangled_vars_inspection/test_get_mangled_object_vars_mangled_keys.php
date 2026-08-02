<?php
// vybe-test: php/php_object_mangled_vars_inspection/test_get_mangled_object_vars_mangled_keys
// origin: languages/php/tests/php/test_php_object_mangled_vars_inspection.rs

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

class Base {
    private int $secret = 42;
}

$vars = get_mangled_object_vars(new Base());
$keys = array_keys($vars);
echo (str_contains($keys[0], 'Base') && str_contains($keys[0], 'secret')) ? 'mangled' : 'plain', "\n";

__vybe_check(ob_get_clean(), "mangled");
