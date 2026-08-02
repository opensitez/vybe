<?php
// vybe-test: php/named_arguments/named_args_instance_method
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

class Formatter {
    public function format(string $value, int $decimals = 2, string $dec_point = '.', string $thousands_sep = ','): string {
        return number_format((float)$value, $decimals, $dec_point, $thousands_sep);
    }
}
$f = new Formatter();
echo $f->format(value: '1234567.891', decimals: 1) . "\n";

__vybe_check(ob_get_clean(), "1,234,567.9");
