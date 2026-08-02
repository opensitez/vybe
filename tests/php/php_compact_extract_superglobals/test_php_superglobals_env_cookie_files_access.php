<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_superglobals_env_cookie_files_access
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs
// vybe-test-mode: compile

$_ENV["APP_KEY"] = "base64:secret";
$_COOKIE["session"] = "cookie_val";
$_FILES["upload"] = ["name" => "photo.jpg", "size" => 1024];

echo $_ENV["APP_KEY"] . " " . $_COOKIE["session"] . " " . $_FILES["upload"]["name"];
