<?php
// vybe-test: php/mb_strings/mb_internal_encoding_set
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$old = mb_internal_encoding();
mb_internal_encoding('UTF-8');
echo mb_internal_encoding() === 'UTF-8' ? 'set to UTF-8' : 'failed';
mb_internal_encoding($old); // restore
