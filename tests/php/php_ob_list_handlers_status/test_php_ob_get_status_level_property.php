<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_level_property
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

ob_start();
$s1 = ob_get_status(false);
ob_start();
$s2 = ob_get_status(false);
ob_end_clean();
ob_end_clean();
echo $s2["level"] === $s1["level"] + 1 ? "STATUS_LEVEL_INC_OK" : "FAIL";
