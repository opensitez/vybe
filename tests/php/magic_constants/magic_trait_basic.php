<?php
// vybe-test: php/magic_constants/magic_trait_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

trait MyTrait {
    public function traitName(): string { return __TRAIT__; }
}
class UsesTrait { use MyTrait; }
echo (new UsesTrait())->traitName();
