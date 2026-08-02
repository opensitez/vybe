<?php
// vybe-test: php/php_ob_implicit_flush_auto/test_php_ob_implicit_flush_default_off
// origin: languages/php/tests/php/test_php_ob_implicit_flush_auto.rs
// vybe-test-mode: compile

ob_implicit_flush(0);
echo "Implicit OFF";
