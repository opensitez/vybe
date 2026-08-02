<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_globals_array_superglobal_access
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs
// vybe-test-mode: compile

$globalVar = "Global Scope";
function testGlobal() {
    echo $GLOBALS["globalVar"];
}
testGlobal();
