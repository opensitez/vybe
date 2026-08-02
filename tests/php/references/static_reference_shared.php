<?php
// vybe-test: php/references/static_reference_shared
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function nextId(): int {
    static $id = 0;
    $id++;
    return $id;
}
$a = nextId();
$b = nextId();
$c = nextId();
echo "$a,$b,$c";
