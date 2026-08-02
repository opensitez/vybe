<?php
// vybe-test: php/modern_php_deep/enum_with_trait
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

trait HasLabel {
    public function label(): string {
        return ucfirst(strtolower($this->name));
    }
}
enum Color {
    use HasLabel;
    case Red;
    case Green;
    case Blue;
}
echo Color::Red->label();
echo Color::Green->label();

__vybe_check(ob_get_clean(), "RedGreen");
