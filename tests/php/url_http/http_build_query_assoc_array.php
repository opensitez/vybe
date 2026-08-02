<?php
// vybe-test: php/url_http/http_build_query_assoc_array
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$params = ['name' => 'Alice', 'age' => 30, 'city' => 'Paris'];
$qs = http_build_query($params);
echo $qs;
