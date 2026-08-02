<?php
// vybe-test: php/namespaces/namespace_current_constant_from_nested_block
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Shop\Models {
    class Product {
        public function scope(): string { return __NAMESPACE__; }
    }
}
namespace App {
    $obj = new \Shop\Models\Product();
    echo $obj->scope();
}
