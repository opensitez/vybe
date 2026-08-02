<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_readonly_class_with_parent_readonly_property
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

readonly class BaseConfig {
    public function __construct(public readonly string $env) {}
}

readonly class ServiceConfig extends BaseConfig {
    public function __construct(string $env, public readonly string $region) {
        parent::__construct($env);
    }
}

$cfg = new ServiceConfig("prod", "eu");
echo $cfg->env . "|" . $cfg->region;

__vybe_check(ob_get_clean(), "prod|eu");
