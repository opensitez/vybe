<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_with_callback_receives_full_chunk_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start(function(string $chunk): string {
    return strtoupper($chunk);
});
echo 'a';
echo 'b';
$c = ob_get_clean();
echo $c;
