<?php
// vybe-test: php/namespaces/namespace_use_import_in_anonymous_function_scope
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Lib {
    function norm(string $v): string { return "[$v]"; }
}
namespace App {
    use function Lib\norm;
    $render = function() use () {
        return norm('x');
    };
    echo $render();
}
