<?php
// vybe-test: php/enum_from_cases/enum_static_method
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

enum Level: int {
    case Low = 1; case High = 2;
    public static function default(): self { return self::Low; }
}
echo Level::default()->name;

__vybe_check(ob_get_clean(), "Low");
