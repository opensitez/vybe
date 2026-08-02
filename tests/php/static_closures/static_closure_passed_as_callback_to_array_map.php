<?php
// vybe-test: php/static_closures/static_closure_passed_as_callback_to_array_map
// origin: languages/php/tests/php/test_static_closures.rs

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

$double = static function(int $n): int { return $n * 2; };
$result = array_map($double, [1, 2, 3]);
echo implode(',', $result);

__vybe_check(ob_get_clean(), "2,4,6");
