<?php
// vybe-test: php/password_security/password_get_info_unknown
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$info = password_get_info('not-a-hash');
echo $info['algo'] === 0 ? 'unknown algo' : 'known algo';
