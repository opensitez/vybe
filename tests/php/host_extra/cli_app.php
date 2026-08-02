<?php
// vybe-test: php/host_extra/cli_app
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

echo "Running on: " . php_uname() . "\n";
echo "User: " . get_current_user() . "\n";
echo "CWD: " . getcwd() . "\n";
echo "PHP: " . phpversion() . "\n";
$name = readline();
echo "Hello, " . $name . "!\n";
