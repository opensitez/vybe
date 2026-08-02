<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_static_targets
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

class DynamicStatic {
    public static function value(int $base, int $delta): int { return $base + $delta; }
}
$method = 'value';
$out = DynamicStatic::$method(10, 5);
echo $out;

__vybe_check(ob_get_clean(), "15");
