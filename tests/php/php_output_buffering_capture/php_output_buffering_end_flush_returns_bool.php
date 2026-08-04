<?php
// vybe-test: php/php_output_buffering_capture/php_output_buffering_end_flush_returns_bool
// origin: languages/php/tests/php/test_php_output_buffering_capture.rs

ob_start(); echo 'ping'; $ok = ob_end_flush(); echo ':' . ($ok ? '1' : '0');
