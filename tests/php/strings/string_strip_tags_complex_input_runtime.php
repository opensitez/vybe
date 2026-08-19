<?php
// vybe-test: php/strings/string_strip_tags_complex_input_runtime
// origin: languages/php/tests/php/test_strings.rs

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

$html = '<div><span>safe</span> &amp; <script>bad()</script></div>';
echo strip_tags($html, '<div><span>');

__vybe_check(ob_get_clean(), "<div><span>safe</span> &amp; bad()</div>");
