<?php
// vybe-test: php/named_arguments/named_args_with_interface_method
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

interface Converter {
    public function convert(float $value, string $from, string $to): float;
}
class TempConverter implements Converter {
    public function convert(float $value, string $from, string $to): float {
        if ($from === 'C' && $to === 'F') return $value * 9/5 + 32;
        if ($from === 'F' && $to === 'C') return ($value - 32) * 5/9;
        return $value;
    }
}
$c = new TempConverter();
echo $c->convert(value: 100.0, from: 'C', to: 'F') . "\n";
echo $c->convert(value: 32.0, from: 'F', to: 'C') . "\n";

__vybe_check(ob_get_clean(), "212\n0");
