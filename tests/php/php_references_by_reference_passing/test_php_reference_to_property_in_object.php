<?php
// vybe-test: php/php_references_by_reference_passing/test_php_reference_to_property_in_object
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs
// vybe-test-mode: compile

class State {
    public string $status = "initial";
}

$s = new State();
$ref = &$s->status;
$ref = "updated";
echo $s->status;
