<?php
// vybe-test: php/named_arguments/named_args_passed_through_wrapper
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

function inner(string $prefix, string $suffix, string $sep = '-'): string {
    return $prefix . $sep . $suffix;
}
function outer(string $prefix, string $suffix, string $sep = '-'): string {
    return inner(prefix: $prefix, suffix: $suffix, sep: $sep);
}
echo outer(prefix: 'foo', suffix: 'bar') . "\n";
echo outer(suffix: 'baz', prefix: 'qux', sep: ':') . "\n";

__vybe_check(ob_get_clean(), "foo-bar\nqux:baz");
