<?php
// vybe-test: php/string_extra_builtins/strtok_tokenize_by_delimiter
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

$token = strtok("Hello World PHP", " ");
$parts = [];
while ($token !== false) {
    $parts[] = $token;
    $token = strtok(" ");
}
echo count($parts);
echo $parts[0];


__vybe_check(ob_get_clean(), "3Hello");
