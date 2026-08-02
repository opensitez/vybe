<?php
// vybe-test: php/namespaces/namespace_dynamic_class_name_with_braces
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Domain {
    class Service { public function label(): string { return 'service'; } }
}
namespace App {
    $c = 'Service';
    $fqcn = "\\Domain\\$c";
    $obj = new $fqcn();
    echo $obj->label();
}
