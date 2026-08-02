<?php
// vybe-test: php/strings/base64
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$e = base64_encode('hello'); echo base64_decode($e);
