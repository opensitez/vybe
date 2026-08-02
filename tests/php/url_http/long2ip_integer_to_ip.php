<?php
// vybe-test: php/url_http/long2ip_integer_to_ip
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$ip = long2ip(2130706433);
echo $ip;
$roundtrip = long2ip(ip2long('10.0.0.1'));
echo $roundtrip;
