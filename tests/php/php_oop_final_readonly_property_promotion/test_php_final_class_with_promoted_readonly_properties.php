<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_final_class_with_promoted_readonly_properties
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs

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

final class ImmutableConfig {
    public function __construct(
        public readonly string $env,
        public readonly bool $debug = false
    ) {}
}

$cfg = new ImmutableConfig("production");
echo "{$cfg->env} debug=" . ($cfg->debug ? "1" : "0");

__vybe_check(ob_get_clean(), "production debug=0");
