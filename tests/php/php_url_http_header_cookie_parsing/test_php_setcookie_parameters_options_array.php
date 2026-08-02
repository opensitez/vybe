<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_setcookie_parameters_options_array
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

if (!headers_sent()) {
    setcookie("session_id", "abc123xyz", [
        "expires" => time() + 3600,
        "path" => "/",
        "domain" => "example.com",
        "secure" => true,
        "httponly" => true,
        "samesite" => "Lax",
    ]);
}
