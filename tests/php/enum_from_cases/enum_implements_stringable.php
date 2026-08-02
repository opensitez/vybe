<?php
// vybe-test: php/enum_from_cases/enum_implements_stringable
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
enum Suit: string implements HasLabel {
    case Hearts = 'hearts';
    public function label(): string { return $this->name . '(' . $this->value . ')'; }
}
echo Suit::Hearts->label();

__vybe_check(ob_get_clean(), "Hearts(hearts)");
