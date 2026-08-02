<?php
// vybe-test: php/generator_errors/generator_from_array_walk
// origin: languages/php/tests/php/test_generator_errors.rs

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

$acc = [];
$gen = (function() use (&$acc) {
    foreach ([1, 2] as $v) { $acc[] = $v; yield $v; }
})();
iterator_to_array($gen);
echo implode('+', $acc);

__vybe_check(ob_get_clean(), "1+2");
