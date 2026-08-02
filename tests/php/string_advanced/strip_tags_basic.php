<?php
// vybe-test: php/string_advanced/strip_tags_basic
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo strip_tags("<p>Hello <b>World</b></p>");
echo "\n";
echo strip_tags("<a href='url'>click</a> here", "<a>");
echo "\n";

__vybe_check(ob_get_clean(), "Hello World\n<a href='url'>click</a> here");
