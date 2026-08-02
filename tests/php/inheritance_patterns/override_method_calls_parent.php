<?php
// vybe-test: php/inheritance_patterns/override_method_calls_parent
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

class Logger {
    public function log(string $msg): string { return "[$msg]"; }
}
class TimestampLogger extends Logger {
    public function log(string $msg): string { return parent::log("2024:" . $msg); }
}
echo (new TimestampLogger)->log('test'), "\n";

__vybe_check(ob_get_clean(), "[2024:test]");
