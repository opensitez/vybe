<?php
// vybe-test: php/namespaces/use_const
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Constants;
const PI = 3.14159;
const E  = 2.71828;

namespace App;
use const Constants\PI;
use const Constants\E;
echo round(PI, 2) . ',' . round(E, 2);
