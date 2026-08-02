<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_with_chunk_size_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start(null, 32, false);
echo str_repeat('a', 5);
$inside = ob_get_contents();
ob_end_clean();
echo $inside . '|';
