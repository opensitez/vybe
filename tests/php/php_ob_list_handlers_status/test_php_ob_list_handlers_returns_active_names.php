<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_list_handlers_returns_active_names
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs

ob_start();
$handlers = ob_list_handlers();
ob_end_clean();

echo "HandlerName: " . ($handlers[0] ?? "");
