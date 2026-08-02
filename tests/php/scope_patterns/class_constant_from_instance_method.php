<?php
// vybe-test: php/scope_patterns/class_constant_from_instance_method
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

class Protocol {
    const VERSION = '1.0';
    public function version(): string { return self::VERSION; }
}
echo (new Protocol())->version();
