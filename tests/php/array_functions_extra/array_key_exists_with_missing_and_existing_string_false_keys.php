<?php
// vybe-test: php/array_functions_extra/array_key_exists_with_missing_and_existing_string_false_keys
// origin: languages/php/tests/php/test_array_functions_extra.rs

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

$a = ['' => 'empty', '0' => 'zero', false => 'false-key'];
echo array_key_exists('', $a) ? 'yes' : 'no';
echo '|';
echo array_key_exists(0, $a) ? 'zero' : 'nozero';
echo '|';
echo array_key_exists('0', $a) ? 'zero2' : 'nozero2';

__vybe_check(ob_get_clean(), "yes|zero|zero2");
