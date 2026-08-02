<?php
// vybe-test: php/output_buffering/ob_flush_without_contents_preserves_outer_output
// origin: languages/php/tests/php/test_output_buffering.rs

echo 'outer-start-';
ob_start();
echo 'inner';
ob_flush();
echo '-outer-end';
ob_end_clean();
