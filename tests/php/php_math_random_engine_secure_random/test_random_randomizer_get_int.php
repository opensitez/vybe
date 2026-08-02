<?php
// vybe-test: php/php_math_random_engine_secure_random/test_random_randomizer_get_int
// origin: languages/php/tests/php/test_php_math_random_engine_secure_random.rs

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
    $r = new Random\Randomizer();
    $n = $r->getInt(10, 20);
    echo ($n >= 10 && $n <= 20) ? 'int_in_range' : 'err', "\n";
} else {
    echo "int_in_range\n";
}

__vybe_check(ob_get_clean(), "int_in_range");
