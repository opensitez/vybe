<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_with_nested_transform_chain_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start(function(string $chunk): string {
    return '[' . $chunk . ']';
});
ob_start(function(string $chunk): string {
    return strtoupper($chunk);
});
echo 'ok';
$inner = ob_get_clean();
echo $inner;
ob_end_flush();
