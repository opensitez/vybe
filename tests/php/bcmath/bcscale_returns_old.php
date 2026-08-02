<?php
// vybe-test: php/bcmath/bcscale_returns_old
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

bcscale(0);
$old = bcscale(5);
echo is_bool($old) || is_int($old) ? 'ok' : 'fail';
bcscale(0);
