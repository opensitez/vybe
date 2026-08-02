<?php
// vybe-test: php/url_http/http_build_query_custom_separator
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$params = ['a' => 1, 'b' => 2, 'c' => 3];
$qs = http_build_query($params, '', '&amp;');
echo $qs;
