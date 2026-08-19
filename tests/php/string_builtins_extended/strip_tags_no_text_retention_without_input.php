<?php
// vybe-test: php/string_builtins_extended/strip_tags_no_text_retention_without_input
// origin: languages/php/tests/php/test_string_builtins_extended.rs

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

echo strip_tags("<script>alert(1)</script>");
echo "|";
echo strip_tags("<b>safe</b><i>ok</i>", "<i>");


__vybe_check(ob_get_clean(), "alert(1)|safe<i>ok</i>");
