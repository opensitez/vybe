<?php
// vybe-test: php/generators_advanced/generator_yield_null_explicitly
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function nulls(int $count) {
    for ($i = 0; $i < $count; $i++) {
        yield null;
    }
}
$c = 0;
foreach (nulls(3) as $v) {
    echo $v === null ? "null" : "not-null";
    $c++;
}
echo $c;

__vybe_check(ob_get_clean(), "nullnullnull3");
