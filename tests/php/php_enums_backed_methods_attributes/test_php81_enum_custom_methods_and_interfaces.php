<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_enum_custom_methods_and_interfaces
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs

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

interface Colorable {
    public function color(): string;
}

enum Priority: int implements Colorable {
    case Low = 1;
    case Medium = 2;
    case High = 3;

    public function color(): string {
        return match($this) {
            self::Low => "green",
            self::Medium => "yellow",
            self::High => "red",
        };
    }
}

echo Priority::High->color();

__vybe_check(ob_get_clean(), "red");
