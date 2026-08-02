<?php
// vybe-test: php/output_buffering/ob_get_contents_after_clean_returns_empty
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'tmp';
ob_clean();
echo ob_get_contents();
ob_end_clean();
echo 'end';
