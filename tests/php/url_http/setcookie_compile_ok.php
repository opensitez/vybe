<?php
// vybe-test: php/url_http/setcookie_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

setcookie('session_id', 'abc123', time() + 3600, '/', '', true, true);
echo 'cookie set';
