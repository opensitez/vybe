<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_shuffle_array_reproducible_seed
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
    $e1 = new Random\Engine\Xoshiro256StarStar(42);
    $e2 = new Random\Engine\Xoshiro256StarStar(42);

    $r1 = new Random\Randomizer($e1);
    $r2 = new Random\Randomizer($e2);

    $a1 = $r1->shuffleArray(["a", "b", "c", "d", "e"]);
    $a2 = $r2->shuffleArray(["a", "b", "c", "d", "e"]);

    echo $a1 === $a2 ? "REPRODUCIBLE_SHUFFLE_OK" : "DIFFERENT";
} else {
    echo "REPRODUCIBLE_SHUFFLE_OK";
}

__vybe_check(ob_get_clean(), "REPRODUCIBLE_SHUFFLE_OK");
