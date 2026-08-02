<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_bcscale_default_setting
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs
// vybe-test-mode: compile

bcscale(4);
echo bcadd("1.11111", "2.22222");
