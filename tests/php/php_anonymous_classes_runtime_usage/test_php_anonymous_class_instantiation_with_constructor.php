<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_instantiation_with_constructor
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs

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

$logger = new class("APP_LOG") {
    public function __construct(public string $prefix) {}
    public function info(string $msg): string {
        return "[{$this->prefix}] $msg";
    }
};

echo $logger->info("Service started");

__vybe_check(ob_get_clean(), "[APP_LOG] Service started");
