<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_outer_scope_capture_via_constructor
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

$config = ["debug" => true, "env" => "staging"];

$service = new class($config) {
    public function __construct(private array $cfg) {}
    public function getEnv(): string { return $this->cfg["env"]; }
};

echo $service->getEnv();

__vybe_check(ob_get_clean(), "staging");
