<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_implementing_interface_with_method
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs

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

interface Labelled {
    public function label(): string;
}

enum Currency: string implements Labelled {
    case USD = "USD";
    case EUR = "EUR";
    case GBP = "GBP";

    public function label(): string {
        return match($this) {
            self::USD => "US Dollar ($)",
            self::EUR => "Euro (€)",
            self::GBP => "Pound Sterling (£)",
        };
    }
}

echo Currency::EUR->label();

__vybe_check(ob_get_clean(), "Euro (€)");
