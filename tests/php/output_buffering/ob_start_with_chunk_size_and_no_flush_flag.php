<?php
// vybe-test: php/output_buffering/ob_start_with_chunk_size_and_no_flush_flag
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(null, 1024, false);
echo 'chunk';
$c = ob_get_contents();
ob_end_clean();
echo $c;
