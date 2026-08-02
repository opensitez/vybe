<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_get_contents_without_clearing
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start();
echo "Persistent";
$c1 = ob_get_contents();
$c2 = ob_get_contents();
ob_end_clean();
echo $c1 === "Persistent" && $c2 === "Persistent" ? "PERSISTENT_CONTENTS_OK" : "FAIL";
