<?php
// vybe-test: php/oop_advanced/mixed_type_in_method_signature
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Converter {
    public function toInt(mixed $value): int {
        return (int) $value;
    }
    public function toBool(mixed $value): bool {
        return (bool) $value;
    }
    public function toStr(mixed $value): string {
        return (string) $value;
    }
}
$c = new Converter();
echo $c->toInt("42"), "\n";
echo $c->toBool(0) ? "true" : "false", "\n";
echo $c->toStr(3.14), "\n";

__vybe_check(ob_get_clean(), "42\nfalse\n3.14");
