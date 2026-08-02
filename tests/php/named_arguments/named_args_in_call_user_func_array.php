<?php
// vybe-test: php/named_arguments/named_args_in_call_user_func_array
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

function buildTag(string $tag, string $content, string $class = ''): string {
    $cls = $class ? " class=\"$class\"" : '';
    return "<$tag$cls>$content</$tag>";
}
$result = call_user_func_array('buildTag', ['tag' => 'div', 'content' => 'Hello', 'class' => 'greeting']);
echo $result . "\n";

__vybe_check(ob_get_clean(), "<div class=\"greeting\">Hello</div>");
