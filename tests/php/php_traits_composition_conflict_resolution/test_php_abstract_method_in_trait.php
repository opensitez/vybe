<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_abstract_method_in_trait
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

trait IdentifiableTrait {
    abstract public function getId(): int;
    
    public function getFormattedId(): string {
        return "ID#" . $this->getId();
    }
}

class Invoice {
    use IdentifiableTrait;
    public function getId(): int { return 42; }
}

$inv = new Invoice();
echo $inv->getFormattedId();
