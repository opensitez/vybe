<?php
// vybe-test: php/strings/urlencode_decode
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$e = urlencode('hello world'); echo urldecode($e);
