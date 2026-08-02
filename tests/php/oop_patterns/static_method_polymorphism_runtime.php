<?php
// vybe-test: php/oop_patterns/static_method_polymorphism_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class FormatterBase {
    public static function supports(): string { return 'base'; }
}
class JsonFormatter extends FormatterBase {
    public static function supports(): string { return 'json'; }
}
class CsvFormatter extends FormatterBase {
    public static function supports(): string { return 'csv'; }
}
echo JsonFormatter::supports();
echo '|' . CsvFormatter::supports();
echo '|' . FormatterBase::supports();

__vybe_check(ob_get_clean(), "json|csv|base");
