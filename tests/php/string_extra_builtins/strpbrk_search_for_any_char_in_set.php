<?php
// vybe-test: php/string_extra_builtins/strpbrk_search_for_any_char_in_set
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

$result = strpbrk("This is a test", "aeiou");
echo is_string($result) ? "found" : "not";
$none = strpbrk("bcdfg", "aeiou");
echo $none === false ? "false" : "found";


__vybe_check(ob_get_clean(), "foundfalse");
