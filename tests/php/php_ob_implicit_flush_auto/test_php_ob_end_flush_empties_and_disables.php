<?php
// vybe-test: php/php_ob_implicit_flush_auto/test_php_ob_end_flush_empties_and_disables
// origin: languages/php/tests/php/test_php_ob_implicit_flush_auto.rs
// vybe-test-mode: compile

ob_start();
echo "Buffered Data";
$levelBefore = ob_get_level();
ob_end_flush();
$levelAfter = ob_get_level();
echo $levelAfter === $levelBefore - 1 ? "END_FLUSH_DECR_LEVEL_OK" : "FAIL";
