<?php
// vybe-test: php/php_references_by_reference_passing/test_php_foreach_by_reference_dangling_reference_fix
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs
// vybe-test-mode: compile

$nums = [1, 2, 3];
foreach ($nums as &$v) {
    $v *= 2;
}
unset($v); // Best practice: unset reference after loop

foreach ($nums as $v) {
    echo $v . "\n";
}
