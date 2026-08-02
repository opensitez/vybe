<?php
// vybe-test: php/output_buffering/ob_callback_exception_passthrough_not_masked
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(function(string $buf): string {
    if ($buf === 'boom') { throw new Exception('cb'); }
    return strtoupper($buf);
});
try {
    echo 'boom';
    ob_get_clean();
    echo 'no';
} catch (Exception $e) {
    echo 'caught';
}
