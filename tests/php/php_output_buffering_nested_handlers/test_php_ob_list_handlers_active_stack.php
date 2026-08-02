<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_list_handlers_active_stack
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
ob_start();
$handlers = ob_list_handlers();
ob_end_clean();
ob_end_clean();
echo implode(", ", $handlers);
