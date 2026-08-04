<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_callback_handler
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(function(string $chunk) { return strtoupper($chunk); }); echo 'ab'; $v = ob_get_clean(); echo $v;
