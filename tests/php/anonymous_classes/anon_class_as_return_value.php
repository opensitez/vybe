<?php
// vybe-test: php/anonymous_classes/anon_class_as_return_value
// origin: languages/php/tests/php/test_anonymous_classes.rs

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

function makeLogger(string $prefix) {
    return new class($prefix) {
        public function __construct(private string $p) {}
        public function log(string $m): string { return $this->p . ': ' . $m; }
    };
}
echo makeLogger('[INFO]')->log('started');

__vybe_check(ob_get_clean(), "[INFO]: started");
