<?php
// vybe-test: php/url_http/http_build_query_numeric_array
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$params = ['foo', 'bar', 'baz'];
$qs = http_build_query($params);
echo $qs;
