<?php
// vybe-test: php/namespaces/enum_in_namespace
// origin: languages/php/tests/php/test_namespaces.rs

namespace namespaces;

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
        return "enum_in_namespace_ok";
    }
}

$g = new Greeter();
echo $g->hello();

__vybe_check(ob_get_clean(), "enum_in_namespace_ok");
