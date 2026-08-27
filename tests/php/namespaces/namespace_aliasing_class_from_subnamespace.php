<?php
// vybe-test: php/namespaces/namespace_aliasing_class_from_subnamespace
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
        return "namespace_aliasing_class_from_subnamespace_ok";
    }
}

$g = new Greeter();
echo $g->hello();

__vybe_check(ob_get_clean(), "namespace_aliasing_class_from_subnamespace_ok");
