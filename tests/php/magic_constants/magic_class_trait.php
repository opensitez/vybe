<?php
// vybe-test: php/magic_constants/magic_class_trait
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

trait Identified {
    public function identify(): string { return __CLASS__; }
}
class Alpha { use Identified; }
class Beta  { use Identified; }
echo (new Alpha())->identify();
echo (new Beta())->identify();
