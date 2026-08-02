<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_get_filters_list
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs
// vybe-test-mode: compile

$filters = stream_get_filters();
echo in_array("string.rot13", $filters) ? "ROT13_FILTER_AVAILABLE" : "NO_ROT13";
