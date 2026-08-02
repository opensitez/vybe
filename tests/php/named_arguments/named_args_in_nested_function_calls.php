<?php
// vybe-test: php/named_arguments/named_args_in_nested_function_calls
// origin: languages/php/tests/php/test_named_arguments.rs

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

function wrap(string $inner, string $outer): string {
    return "<$outer>$inner</$outer>";
}
function buildHtml(string $text, string $inner_tag = 'span', string $outer_tag = 'div'): string {
    return wrap(inner: "<$inner_tag>$text</$inner_tag>", outer: $outer_tag);
}
echo buildHtml(text: 'hello', inner_tag: 'b') . "\n";

__vybe_check(ob_get_clean(), "<div><b>hello</b></div>");
