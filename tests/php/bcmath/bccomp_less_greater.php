<?php
// vybe-test: php/bcmath/bccomp_less_greater
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

echo bccomp('1', '2');    // -1
echo bccomp('2', '1');    // 1
echo bccomp('1.5', '1.6', 1); // -1
