<?php
// vybe-test: php/url_http/urlencode_vs_rawurlencode
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$str = 'hello world & more';
echo urlencode($str);
echo rawurlencode($str);
