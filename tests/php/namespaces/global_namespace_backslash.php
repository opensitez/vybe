<?php
// vybe-test: php/namespaces/global_namespace_backslash
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App;
$len = \strlen("hello");
echo $len;
