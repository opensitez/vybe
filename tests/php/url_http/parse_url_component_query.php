<?php
// vybe-test: php/url_http/parse_url_component_query
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$query = parse_url('https://example.com/search?q=hello&page=2', PHP_URL_QUERY);
echo $query;
