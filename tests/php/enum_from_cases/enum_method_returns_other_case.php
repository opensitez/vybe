<?php
// vybe-test: php/enum_from_cases/enum_method_returns_other_case
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

enum Toggle { case On; case Off;
    public function flip(): self {
        return match($this) { Toggle::On => Toggle::Off, Toggle::Off => Toggle::On };
    }
}
echo Toggle::On->flip()->name;

__vybe_check(ob_get_clean(), "Off");
