<?php
// vybe-test: php/string_extra_builtins/parse_str_query_string_into_variables
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

parse_str("name=Alice&age=30&city=Paris", $output);
echo $output["name"];
echo $output["age"];
echo $output["city"];
