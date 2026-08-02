<?php
// vybe-test: php/output_buffering/ob_get_contents_after_flush_is_empty
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'hi';
ob_flush();
echo ob_get_contents();
ob_end_clean();
