<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_get_level_after_nested_failures
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); $before = ob_get_level(); ob_start(); echo 'x'; ob_end_flush(); $after = ob_get_level(); ob_end_clean(); echo $before . ':' . $after;
