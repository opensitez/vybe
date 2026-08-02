<?php
// vybe-test: php/namespaces/namespace_fully_qualified_calls_from_root
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App {
    function tag(string $v): string { return "app:$v"; }
}
function tag(string $v): string { return "global:$v"; }
echo \App\tag('x') === 'app:x' ? 'app' : 'no';
echo '|' . tag('y') === 'global:y' ? 'global' : 'noglobal';
