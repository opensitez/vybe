<?php
// vybe-test: php/enums_advanced/enum_static_method
// origin: languages/php/tests/php/test_enums_advanced.rs

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

enum Color: string {
    case Red = 'red'; case Green = 'green'; case Blue = 'blue';
    public static function fromHex(string $hex): self {
        return match($hex) { '#ff0000' => self::Red, '#00ff00' => self::Green, default => self::Blue };
    }
}
echo Color::fromHex('#ff0000')->value;

__vybe_check(ob_get_clean(), "red");
