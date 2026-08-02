<?php
// vybe-test: php/url_http/parse_url_component_path
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$path = parse_url('https://example.com/foo/bar?x=1', PHP_URL_PATH);
echo $path;
