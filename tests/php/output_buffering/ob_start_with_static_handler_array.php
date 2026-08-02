<?php
// vybe-test: php/output_buffering/ob_start_with_static_handler_array
// origin: languages/php/tests/php/test_output_buffering.rs

class BufferFilters {
    public static function frame(string $buf): string { return '[' . $buf . ']'; }
}
ob_start([BufferFilters::class, 'frame']);
echo 'payload';
$out = ob_get_clean();
echo $out;
