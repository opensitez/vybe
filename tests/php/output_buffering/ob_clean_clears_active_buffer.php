<?php
// vybe-test: php/output_buffering/ob_clean_clears_active_buffer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'remove';
ob_clean();
echo 'kept';
$c = ob_get_clean();
echo '|' . $c;
