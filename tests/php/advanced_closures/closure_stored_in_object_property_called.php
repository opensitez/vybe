<?php
// vybe-test: php/advanced_closures/closure_stored_in_object_property_called
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class Handler {
    public Closure $onEvent;
    public function __construct() {
        $this->onEvent = function(string $name): string { return 'handled: ' . $name; };
    }
}
$h = new Handler();
echo ($h->onEvent)('click');
