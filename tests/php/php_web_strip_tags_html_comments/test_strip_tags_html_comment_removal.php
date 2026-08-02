<?php
// vybe-test: php/php_web_strip_tags_html_comments/test_strip_tags_html_comment_removal
// origin: languages/php/tests/php/test_php_web_strip_tags_html_comments.rs

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

$html = "<!-- secret comment -->Hello <b>World</b>";
echo strip_tags($html), "\n";

__vybe_check(ob_get_clean(), "Hello World");
