<?php
// vybe-test: php/url_http/ip2long_ipv4_address
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$n = ip2long('192.168.1.1');
echo $n > 0 ? 'positive' : 'fail';
$loopback = ip2long('127.0.0.1');
echo $loopback;
