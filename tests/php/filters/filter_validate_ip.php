<?php
// vybe-test: php/filters/filter_validate_ip
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$ips = ['192.168.1.1', '256.0.0.1', '::1', '2001:db8::1', 'not-an-ip'];
foreach ($ips as $ip) {
    echo filter_var($ip, FILTER_VALIDATE_IP) !== false ? 'valid' : 'invalid';
    echo ' ';
}
