<?php
// vybe-test: php/output_buffering/ob_start_with_output_handler_array
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(new class {
    public function __invoke(string $buf): string {
        return str_replace('a', 'A', $buf);
    }
});
echo 'java';
echo ob_get_clean();
