<?php
// vybe-test: php/url_http/inet_pton_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$packed = inet_pton('127.0.0.1');
echo $packed !== false ? 'ok' : 'fail';
$packed6 = inet_pton('::1');
echo $packed6 !== false ? 'ok' : 'fail';
