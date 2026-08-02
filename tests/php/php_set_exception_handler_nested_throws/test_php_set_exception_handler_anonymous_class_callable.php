<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_anonymous_class_callable
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$invoked = false;
$handler = new class(&$invoked) {
    private $invoked;
    public function __construct(&$invoked) { $this->invoked = &$invoked; }
    public function __invoke(Throwable $e) { $this->invoked = true; }
};
set_exception_handler($handler);
try {
    throw new Exception("Anon class exception");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $invoked ? "ANON_CLASS_EXCEPTION_HANDLER_OK" : "FAIL";
