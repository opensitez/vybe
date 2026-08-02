<?php
// vybe-test: php/bcmath/bccomp_equal
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

echo bccomp('1.0', '1.00', 2);
echo bccomp('999', '999');
