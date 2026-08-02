<?php
// vybe-test: php/generators/generator_yield_key_value_pairs
// origin: languages/php/tests/php/test_generators.rs

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

function gen(): Generator { yield 'a' => 1; yield 'b' => 2; }
echo json_encode(iterator_to_array(gen()));

__vybe_check(ob_get_clean(), "{\"a\":1,\"b\":2}");
