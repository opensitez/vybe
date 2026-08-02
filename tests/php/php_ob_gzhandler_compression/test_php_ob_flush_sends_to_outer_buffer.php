<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_flush_sends_to_outer_buffer
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start(); // Outer
ob_start(); // Inner
echo "InnerPayload";
ob_flush(); // Flushes inner to outer
$innerRemaining = ob_get_clean();
$outerPayload = ob_get_clean();
echo $innerRemaining === "" && $outerPayload === "InnerPayload" ? "OB_FLUSH_OUTER_OK" : "FAIL";
