<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_final_hook
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs
// vybe-test-mode: compile

class SecureModel {
    public string $id {
        final get => "SECURE_ID";
    }
}

$sm = new SecureModel();
echo $sm->id;
