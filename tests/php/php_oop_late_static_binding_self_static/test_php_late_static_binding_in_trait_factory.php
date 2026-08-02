<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_late_static_binding_in_trait_factory
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs

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

trait FactoryTrait {
    public static function make(string $id): static {
        return new static($id);
    }
}

class Service {
    use FactoryTrait;
    public function __construct(public string $id) {}
}

class BillingService extends Service {}

$service = BillingService::make("billing");
echo get_class($service) . "|" . $service->id;

__vybe_check(ob_get_clean(), "BillingService|billing");
