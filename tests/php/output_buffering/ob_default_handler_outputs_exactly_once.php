<?php
// vybe-test: php/output_buffering/ob_default_handler_outputs_exactly_once
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'one';
echo 'two';
echo 'three';
echo ob_get_clean();
