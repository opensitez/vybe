<?php
// vybe-test: php/url_http/urldecode_vs_rawurldecode
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$encoded = 'hello+world%20%26%20more';
echo urldecode($encoded);
echo rawurldecode($encoded);
