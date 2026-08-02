<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_abstract_protected_method_override
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

abstract class DataProcessor {
    abstract protected function transform(array $data): array;
    
    public function process(array $data): array {
        return $this->transform($data);
    }
}

class CSVProcessor extends DataProcessor {
    public function transform(array $data): array {
        return array_map('strtoupper', $data);
    }
}

$p = new CSVProcessor();
print_r($p->process(["a", "b"]));
