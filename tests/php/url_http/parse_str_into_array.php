<?php
// vybe-test: php/url_http/parse_str_into_array
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

parse_str('name=Bob&score=100&active=1', $out);
echo $out['name'];
echo $out['score'];
echo $out['active'];
