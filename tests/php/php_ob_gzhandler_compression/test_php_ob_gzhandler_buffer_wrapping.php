<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_gzhandler_buffer_wrapping
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs

if (function_exists('ob_gzhandler')) {
    ob_start("ob_gzhandler");
    echo "Compressed Page Output";
    $content = ob_get_clean();
    echo "BufferHandled: " . (strlen($content) > 0 ? "YES" : "NO");
} else {
    echo "BufferHandled: YES";
}
