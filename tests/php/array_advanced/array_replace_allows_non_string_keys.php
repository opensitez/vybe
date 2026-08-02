<?php
// vybe-test: php/array_advanced/array_replace_allows_non_string_keys
// origin: languages/php/tests/php/test_array_advanced.rs

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

$base = [1 => "one", "2" => "string-two", 3.0 => "float-three"];
$patch = [true => "bool-key", 2 => "int-two"];
$result = array_replace($base, $patch);
echo $result[1];
echo "|";
echo $result[2];
echo "|";
echo $result[3];

__vybe_check(ob_get_clean(), "bool-key|int-two|float-three");
