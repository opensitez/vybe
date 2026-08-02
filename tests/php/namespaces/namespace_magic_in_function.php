<?php
// vybe-test: php/namespaces/namespace_magic_in_function
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App\Core;
function getNamespace(): string { return __NAMESPACE__; }
echo getNamespace();
