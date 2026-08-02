<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_exception_hierarchy_instanceof_check
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

$e = new OutOfBoundsException("Index out of bounds");
echo ($e instanceof RuntimeException ? "RUNTIME_EX" : "NO");
echo ($e instanceof Exception ? " EXCEPTION" : " NO");
echo ($e instanceof Throwable ? " THROWABLE" : " NO");
