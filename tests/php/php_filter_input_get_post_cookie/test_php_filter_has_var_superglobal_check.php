<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_has_var_superglobal_check
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_GET["param"] = "value";
echo filter_has_var(INPUT_GET, "param") ? "HAS_VAR" : "NO_VAR";
