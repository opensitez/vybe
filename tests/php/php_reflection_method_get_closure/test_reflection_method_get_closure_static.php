<?php
// vybe-test: php/php_reflection_method_get_closure/test_reflection_method_get_closure_static
// origin: languages/php/tests/php/test_php_reflection_method_get_closure.rs

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

class MathUtil {
    public static function square(int $n): int { return $n * $n; }
}
$rm = new ReflectionMethod(MathUtil::class, 'square');
$closure = $rm->getClosure();
echo $closure(7), "\n";

__vybe_check(ob_get_clean(), "49");
