<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_ob_get_contents_and_clean
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); echo 'hello'; $inner = ob_get_contents(); ob_clean(); echo $inner; ob_end_flush();
