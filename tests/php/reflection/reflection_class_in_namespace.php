<?php
// vybe-test: php/reflection/reflection_class_in_namespace
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

namespace App\Services;
class UserProcessor {}
$rc = new \ReflectionClass(UserProcessor::class);
echo $rc->inNamespace() ? 'in_ns' : 'no';
echo ':' . $rc->getNamespaceName();
