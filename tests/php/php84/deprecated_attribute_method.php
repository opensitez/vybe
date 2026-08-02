<?php
// vybe-test: php/php84/deprecated_attribute_method
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Api {
    #[\Deprecated('Use v2() instead')]
    public function v1(): string { return 'v1'; }
    public function v2(): string { return 'v2'; }
}
echo (new Api())->v2();
