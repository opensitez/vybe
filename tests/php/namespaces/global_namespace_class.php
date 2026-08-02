<?php
// vybe-test: php/namespaces/global_namespace_class
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App;
$e = new \Exception("from global ns");
echo $e->getMessage();
