<?php
// vybe-test: php/php_ob_implicit_flush_auto/test_php_ob_clean_on_inactive_buffer_returns_false
// origin: languages/php/tests/php/test_php_ob_implicit_flush_auto.rs
// vybe-test-mode: compile

while (ob_get_level() > 0) ob_end_clean();
$res = @ob_clean();
echo $res === false ? "INACTIVE_CLEAN_FALSE" : "FAIL";
