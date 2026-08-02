<?php
// vybe-test: php/output_buffering/ob_start_with_numeric_chunk_and_non_flush_flag
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(null, 1, false);
echo 'x';
echo ob_get_length();
echo '|';
echo ob_get_flush() === false ? 'closed' : 'not';
ob_end_clean();
