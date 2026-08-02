<?php
// vybe-test: php/bcmath/bcsub_negative_result
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

echo bcsub('3', '10');
echo bcsub('3.5', '10.2', 1);
