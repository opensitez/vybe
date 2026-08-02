<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_start_chunk_size_trigger
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs

$chunks = 0;
ob_start(function($buffer) use (&$chunks) {
    $chunks++;
    return $buffer;
}, 10); // Chunk size 10 bytes

echo "12345678901"; // 11 bytes triggers chunk handler
$count = $chunks;
ob_end_clean();

echo "ChunksTriggered: " . ($count > 0 ? "YES" : "NO");
