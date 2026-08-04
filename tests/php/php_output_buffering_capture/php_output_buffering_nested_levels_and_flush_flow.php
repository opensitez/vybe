<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_nested_levels_and_flush_flow
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); echo 'a'; ob_start(); echo 'b'; $l1 = ob_get_level(); $inner = ob_get_clean(); echo $l1 . ':' . $inner . ':' . ob_get_contents(); ob_end_flush();
