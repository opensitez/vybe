<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_ip_address_flags
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_SERVER["REMOTE_ADDR"] = "192.168.1.1";
$ip = filter_input(INPUT_SERVER, "REMOTE_ADDR", FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE);
echo $ip === false ? "PRIVATE_IP_FILTERED" : "PUBLIC_IP";
