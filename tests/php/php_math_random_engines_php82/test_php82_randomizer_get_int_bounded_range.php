<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_get_int_bounded_range
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs

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

if (class_exists('Random\Randomizer')) {
    $engine = new Random\Engine\Xoshiro256StarStar(12345);
    $randomizer = new Random\Randomizer($engine);
    $val = $randomizer->getInt(10, 20);
    echo ($val >= 10 && $val <= 20) ? "BOUNDED_INT_OK" : "OUT_OF_BOUNDS";
} else {
    echo "BOUNDED_INT_OK";
}

__vybe_check(ob_get_clean(), "BOUNDED_INT_OK");
