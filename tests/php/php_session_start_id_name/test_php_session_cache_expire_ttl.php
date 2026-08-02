<?php
// vybe-test: php/php_session_start_id_name/test_php_session_cache_expire_ttl
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

@session_cache_expire(30);
echo @session_cache_expire() === 30 ? "CACHE_EXPIRE_OK" : "FAIL";
