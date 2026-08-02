<?php
// vybe-test: php/output_buffering/ob_callback_add_suffix_and_trim
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => trim($buf) . '|ok');
echo '  value  ';
echo ob_get_clean();
