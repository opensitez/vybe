<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_parent_static_method_chaining
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

class ParentFactory {
    public static function boot() { echo "ParentBoot "; }
}

class ChildFactory extends ParentFactory {
    public static function boot() {
        parent::boot();
        echo "ChildBoot";
    }
}

ChildFactory::boot();
