<?php
// vybe-test: php/filters/filter_validate_ip_v6_only
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

echo filter_var('::1',         FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) !== false ? 'v6' : 'fail';
echo filter_var('192.168.1.1', FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) !== false ? 'v6' : 'fail';
