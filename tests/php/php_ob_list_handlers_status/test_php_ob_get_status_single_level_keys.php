<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_single_level_keys
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs

ob_start();
$status = ob_get_status(false);
ob_end_clean();

$keys = ["name", "type", "flags", "level", "chunk_size", "buffer_size", "buffer_used"];
$hasKeys = true;
foreach ($keys as $k) {
    if (!array_key_exists($k, $status)) { $hasKeys = false; break; }
}
echo $hasKeys ? "STATUS_KEYS_OK" : "MISSING_KEYS";
