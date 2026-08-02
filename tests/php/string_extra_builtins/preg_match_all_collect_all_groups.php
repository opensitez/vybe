<?php
// vybe-test: php/string_extra_builtins/preg_match_all_collect_all_groups
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$html = '<a href="http://example.com">one</a> <a href="http://test.org">two</a>';
$count = preg_match_all('/<a href="([^"]+)">([^<]+)<\/a>/', $html, $matches);
echo $count;
echo count($matches[1]);
echo $matches[2][0];
