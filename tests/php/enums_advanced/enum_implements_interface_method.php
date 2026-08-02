<?php
// vybe-test: php/enums_advanced/enum_implements_interface_method
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

interface Colorable { public function hex(): string; }
enum Palette: string implements Colorable {
    case Red = 'red'; case Blue = 'blue';
    public function hex(): string { return match($this) { self::Red => '#FF0000', self::Blue => '#0000FF' }; }
}
echo Palette::Red->hex();

__vybe_check(ob_get_clean(), "#FF0000");
