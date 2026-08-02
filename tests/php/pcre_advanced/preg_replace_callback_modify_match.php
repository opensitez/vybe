<?php
// vybe-test: php/pcre_advanced/preg_replace_callback_modify_match
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$result = preg_replace_callback('/\b(\w)(\w+)\b/', function($m) {
    return strtoupper($m[1]) . strtolower($m[2]);
}, 'hello world from php');
echo $result;
