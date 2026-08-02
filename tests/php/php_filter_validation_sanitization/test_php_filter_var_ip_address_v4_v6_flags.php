<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_ip_address_v4_v6_flags
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$ipv4 = "192.168.1.1";
$ipv6 = "2001:0db8:85a3:0000:0000:8a2e:0370:7334";

echo filter_var($ipv4, FILTER_VALIDATE_IP, FILTER_FLAG_IPV4) ? "V4_OK" : "V4_FAIL";
echo filter_var($ipv6, FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) ? "V6_OK" : "V6_FAIL";
