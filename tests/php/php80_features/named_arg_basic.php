<?php
// vybe-test: php/php80_features/named_arg_basic
// origin: languages/php/tests/php/test_php80_features.rs

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

function createTag(string $tag, string $content, string $class = ''): string {
    $cls = $class ? " class=\"$class\"" : '';
    return "<$tag$cls>$content</$tag>";
}
echo createTag(content: 'Hello', tag: 'p', class: 'greeting');

__vybe_check(ob_get_clean(), "<p class=\"greeting\">Hello</p>");
