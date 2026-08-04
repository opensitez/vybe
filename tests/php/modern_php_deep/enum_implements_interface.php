<?php
// vybe-test: php/modern_php_deep/enum_implements_interface
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

interface HasSymbol {
    public function symbol(): string;
}
enum Currency: string implements HasSymbol {
    case USD = "usd";
    case EUR = "eur";
    case GBP = "gbp";
    public function symbol(): string {
        return match($this) {
            self::USD => "$",
            self::EUR => "€",
            self::GBP => "£" };
    }
}
echo Currency::USD->symbol();
echo Currency::EUR->symbol();
echo Currency::GBP->value;

__vybe_check(ob_get_clean(), "\$€gbp");
