<?php
// vybe-test: php/php_ob_implicit_flush_auto/test_php_ob_get_flush_returns_content_and_flushes
// origin: languages/php/tests/php/test_php_ob_implicit_flush_auto.rs
// vybe-test-mode: compile

ob_start();
echo "Flush Content";
$content = ob_get_flush();
echo $content === "Flush Content" ? "GET_FLUSH_CONTENT_OK" : "FAIL";
