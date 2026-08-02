<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_end_flush_emits_to_outer_buffer
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start(); // Outer
echo "OUTER_START ";
ob_start(); // Inner
echo "INNER_TEXT ";
ob_end_flush(); // Flushes Inner into Outer
echo "OUTER_END";
$final = ob_get_clean();
echo $final;
