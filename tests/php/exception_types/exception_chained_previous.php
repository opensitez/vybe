<?php
// vybe-test: php/exception_types/exception_chained_previous
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

$cause = new RuntimeException('disk full');
$e = new Exception('write failed', 0, $cause);
echo $e->getPrevious()->getMessage();
