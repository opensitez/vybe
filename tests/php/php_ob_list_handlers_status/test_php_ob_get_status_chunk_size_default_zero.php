<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_chunk_size_default_zero
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

ob_start();
$status = ob_get_status(false);
ob_end_clean();
echo isset($status["chunk_size"]) ? "CHUNK_SIZE_KEY_OK" : "FAIL";
