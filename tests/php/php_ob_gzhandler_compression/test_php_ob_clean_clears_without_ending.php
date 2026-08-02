<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_clean_clears_without_ending
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start();
echo "First";
ob_clean();
echo "Second";
$out = ob_get_clean();
echo $out === "Second" ? "OB_CLEAN_OK" : "FAIL";
