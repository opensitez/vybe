<?php
// vybe-test: php/output_buffering/ob_get_flush_applies_callback
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => $buf . '-C');
echo 'x';
$flushed = ob_get_flush();
if ($flushed !== false) {
    echo $flushed;
}
