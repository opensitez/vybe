<?php
// vybe-test: php/filters/filter_sanitize_email
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$email = "user name@exa mple.com";
$clean = filter_var($email, FILTER_SANITIZE_EMAIL);
echo $clean;
