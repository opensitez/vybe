<?php
// vybe-test: php/namespaces/namespace_same_class_name_different_ns
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace V1;
class Response { public function status(): int { return 200; } }

namespace V2;
class Response { public function status(): int { return 201; } }

namespace App;
$r1 = new \V1\Response();
$r2 = new \V2\Response();
echo $r1->status() . ',' . $r2->status();
