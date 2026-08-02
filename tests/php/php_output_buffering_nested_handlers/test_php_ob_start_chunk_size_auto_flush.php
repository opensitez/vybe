<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_chunk_size_auto_flush
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs
// vybe-test-mode: compile

$flushedChunks = [];
ob_start(function($buffer) use (&$flushedChunks) {
    $flushedChunks[] = $buffer;
    return $buffer;
}, chunk_size: 10);

echo "1234567890"; // Should trigger chunk flush
echo "abcdefghij";
ob_end_clean();
