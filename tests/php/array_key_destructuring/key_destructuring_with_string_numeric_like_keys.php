<?php
// vybe-test: php/array_key_destructuring/key_destructuring_with_string_numeric_like_keys
// origin: languages/php/tests/php/test_array_key_destructuring.rs

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

$data = ["0" => "zero", 1 => "one", "2" => "two", "name" => "n"];
["0" => $a, 1 => $b, "2" => $c] = $data;
echo "$a|$b|$c|{$data['name']}";

__vybe_check(ob_get_clean(), "zero|one|two|n");
