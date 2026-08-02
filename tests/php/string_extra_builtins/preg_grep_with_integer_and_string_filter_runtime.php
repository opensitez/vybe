<?php
// vybe-test: php/string_extra_builtins/preg_grep_with_integer_and_string_filter_runtime
// origin: languages/php/tests/php/test_string_extra_builtins.rs

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

$values = ["a1", "b2", "a3", "c4"];
$matches = preg_grep('/^a/', $values);
echo count($matches);
echo '|';
echo implode(',', $matches);

__vybe_check(ob_get_clean(), "2|a1,a3");
