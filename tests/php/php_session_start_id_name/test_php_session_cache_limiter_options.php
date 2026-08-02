<?php
// vybe-test: php/php_session_start_id_name/test_php_session_cache_limiter_options
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

@session_cache_limiter("private");
echo @session_cache_limiter() === "private" ? "CACHE_LIMITER_OK" : "FAIL";
