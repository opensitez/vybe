<?php
// vybe-test: php/programs/fizzbuzz_1_to_15
// origin: languages/php/tests/php/test_programs.rs

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

for ($i = 1; $i <= 15; $i++) {
    if ($i % 15 === 0) echo "FizzBuzz\n";
    elseif ($i % 3 === 0) echo "Fizz\n";
    elseif ($i % 5 === 0) echo "Buzz\n";
    else echo $i . "\n";
}

__vybe_check(ob_get_clean(), "1\n2\nFizz\n4\nBuzz\nFizz\n7\n8\nFizz\nBuzz\n11\nFizz\n13\n14\nFizzBuzz");
