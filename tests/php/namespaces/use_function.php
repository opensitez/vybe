<?php
// vybe-test: php/namespaces/use_function
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Helpers;
function slugify(string $s): string {
    return strtolower(str_replace(' ', '-', $s));
}

namespace App;
use function Helpers\slugify;
echo slugify('Hello World');
