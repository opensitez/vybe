<?php
// vybe-test: php/exception_types/exception_with_message_and_code
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

$e = new Exception('not found', 404);
echo $e->getMessage();
echo $e->getCode();
