<?php
// vybe-test: php/first_class_callables/static_method_first_class_callable
// origin: languages/php/tests/php/test_first_class_callables.rs

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

class MathUtils {
    public static function square(int $n): int { return $n * $n; }
    public static function cube(int $n): int { return $n * $n * $n; }
}
$sq = MathUtils::square(...);
$cu = MathUtils::cube(...);
echo $sq(4) . "\n";
echo $cu(3) . "\n";
$result = array_map($sq, [1, 2, 3, 4]);
echo implode(',', $result) . "\n";

__vybe_check(ob_get_clean(), "16\n27\n1,4,9,16");
