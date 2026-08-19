<?php
// vybe-test: php/oop_patterns/interface_dispatch_with_multiple_implementers_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

interface Formatter {
    public function format(string $value): string;
}

class Upper implements Formatter {
    public function format(string $value): string { return strtoupper($value); }
}

class Lower implements Formatter {
    public function format(string $value): string { return strtolower($value); }
}

function apply_formatter(Formatter $formatter, string $value): string {
    return $formatter->format($value);
}

echo apply_formatter(new Upper(), 'ab');
echo apply_formatter(new Lower(), 'AB');

__vybe_check(ob_get_clean(), "ABab");
