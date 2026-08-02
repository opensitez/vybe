<?php
// vybe-test: php/references/global_reference
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$globalVal = 0;
function modifyGlobal() {
    global $globalVal;
    $globalVal = 42;
}
modifyGlobal();
echo $globalVal;
