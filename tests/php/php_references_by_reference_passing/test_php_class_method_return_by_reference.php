<?php
// vybe-test: php/php_references_by_reference_passing/test_php_class_method_return_by_reference
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs
// vybe-test-mode: compile

class DataContainer {
    private int $value = 42;
    public function &getValue(): int {
        return $this->value;
    }
}

$dc = new DataContainer();
$val = &$dc->getValue();
$val = 100;
echo $dc->getValue();
