<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_list_handlers_order
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); echo 'A'; ob_start(); echo 'B'; $handlers = ob_list_handlers(); $size = count($handlers); ob_end_flush(); ob_end_flush(); echo $size;
