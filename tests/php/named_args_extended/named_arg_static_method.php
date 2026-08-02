<?php
// vybe-test: php/named_args_extended/named_arg_static_method
// origin: languages/php/tests/php/test_named_args_extended.rs

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

class Converter {
    public static function convert(float $value, string $from = 'C', string $to = 'F'): float {
        if ($from === 'C' && $to === 'F') return $value * 9/5 + 32;
        if ($from === 'F' && $to === 'C') return ($value - 32) * 5/9;
        return $value;
    }
}
echo Converter::convert(value: 100.0, to: 'F', from: 'C');

__vybe_check(ob_get_clean(), "212");
