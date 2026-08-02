<?php
// vybe-test: php/type_hints_advanced/nullsafe_static_method
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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

class Factory { public static function create(): ?self { return new self; } public function value(): int { return 42; } }
echo (Factory::create())?->value() ?? 0;

__vybe_check(ob_get_clean(), "42");
