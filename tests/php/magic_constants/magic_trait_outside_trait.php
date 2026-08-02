<?php
// vybe-test: php/magic_constants/magic_trait_outside_trait
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Plain {
    public function trait(): string { return __TRAIT__; }
}
echo (new Plain())->trait() === '' ? 'empty outside trait' : 'has value';
