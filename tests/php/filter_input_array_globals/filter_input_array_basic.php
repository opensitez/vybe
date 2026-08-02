<?php
// vybe-test: php/filter_input_array_globals/filter_input_array_basic
// origin: languages/php/tests/php/test_filter_input_array_globals.rs

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

$_GET = ['name' => 'John', 'age' => '30'];
$args = [
    'name' => FILTER_SANITIZE_STRING,
    'age'  => FILTER_VALIDATE_INT,
];
$res = filter_var_array($_GET, $args);
echo $res['name'] . "|" . $res['age'];

__vybe_check(ob_get_clean(), "John|30");
