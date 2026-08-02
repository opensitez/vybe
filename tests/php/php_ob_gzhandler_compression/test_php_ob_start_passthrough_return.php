<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_start_passthrough_return
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start(fn($s) => strtoupper($s));
echo "lowercase";
$out = ob_get_clean();
echo $out === "LOWERCASE" ? "PASSTHROUGH_STRTOUPPER_OK" : "FAIL";
