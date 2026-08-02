<?php
// vybe-test: php/enum_from_cases/enum_implements_interface_with_method
// origin: languages/php/tests/php/test_enum_from_cases.rs

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

interface HasLabel {
    public function label(): string;
}
enum Color: string implements HasLabel {
    case Red = 'red';
    case Blue = 'blue';
    public function label(): string { return ucfirst($this->value); }
}
echo Color::Red->label();

__vybe_check(ob_get_clean(), "Red");
