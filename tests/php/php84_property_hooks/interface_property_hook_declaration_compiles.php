<?php
// vybe-test: php/php84_property_hooks/interface_property_hook_declaration_compiles
// origin: languages/php/tests/php/test_php84_property_hooks.rs
// vybe-test-mode: compile

interface Shape {
    public float $area { get; }
}
