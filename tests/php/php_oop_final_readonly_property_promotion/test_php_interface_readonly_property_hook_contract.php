<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_interface_readonly_property_hook_contract
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

interface Identifiable {
    public int $id { get; }
}

class Record implements Identifiable {
    public function __construct(public readonly int $id) {}
}

$r = new Record(101);
echo $r->id;
