<?php
// vybe-test: php/namespaces/namespace_fully_qualified_call
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Helpers;
function double(int $n): int { return $n * 2; }

namespace App;
$result = \Helpers\double(21);
echo $result;
