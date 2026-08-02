<?php
// vybe-test: php/bcmath/bcscale_global
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

bcscale(4);
echo bcadd('1', '2');
echo bcdiv('1', '3');
bcscale(0); // reset
