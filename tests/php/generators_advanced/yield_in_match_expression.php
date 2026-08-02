<?php
// vybe-test: php/generators_advanced/yield_in_match_expression
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

function classifyNumbers(array $nums) {
    foreach ($nums as $n) {
        yield match(true) {
            $n < 0  => "neg",
            $n === 0 => "zero",
            $n > 0  => "pos",
        };
    }
}
$result = [];
foreach (classifyNumbers([-2, 0, 3, -1, 5]) as $label) {
    $result[] = $label;
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "neg,zero,pos,neg,pos");
