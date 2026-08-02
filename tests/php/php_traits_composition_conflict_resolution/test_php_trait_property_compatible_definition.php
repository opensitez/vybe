<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_property_compatible_definition
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

trait Configurable {
    public array $options = [];
}

class Settings {
    use Configurable;
    public array $options = []; // Compatible property redeclaration allowed in PHP 8.0+
}

$s = new Settings();
print_r($s->options);
