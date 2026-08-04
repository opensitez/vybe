<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_clean_vs_end_clean_semantics
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); echo 'kept'; ob_clean(); echo 'x'; $level = ob_get_level(); ob_end_clean(); echo $level;
