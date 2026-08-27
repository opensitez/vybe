<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_relative_resolution_with_class_exists
// origin: languages/php/tests/php/test_php_namespaces_runtime.rs

namespace php_namespaces_runtime;

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new \Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

class Greeter {
    public function hello(): string {
        return "php_namespace_relative_resolution_with_class_exists_ok";
    }
}

$g = new Greeter();
echo $g->hello();

__vybe_check(ob_get_clean(), "php_namespace_relative_resolution_with_class_exists_ok");
