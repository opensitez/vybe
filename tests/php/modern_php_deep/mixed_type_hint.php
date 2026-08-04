<?php
// vybe-test: php/modern_php_deep/mixed_type_hint
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function display(mixed $value): string {
    return match(gettype($value)) {
        "integer" => "int:$value",
        "string"  => "str:$value",
        "array"   => "arr:" . count($value),
        "NULL"    => "null",
        default   => "other" };
}
echo display(42);
echo display("hello");
echo display([1, 2, 3]);
echo display(null);

__vybe_check(ob_get_clean(), "int:42str:helloarr:3null");
