<?php
// vybe-test: php/error_handling_deep/error_handler_class
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

class ErrorCollector {
    private array $errors = [];
    public function handle(int $errno, string $errstr): bool {
        $this->errors[] = ['no' => $errno, 'msg' => $errstr];
        return true;
    }
    public function getErrors(): array { return $this->errors; }
    public function count(): int { return count($this->errors); }
}
$collector = new ErrorCollector();
set_error_handler([$collector, 'handle']);
trigger_error("error one", E_USER_NOTICE);
trigger_error("error two", E_USER_WARNING);
restore_error_handler();
echo $collector->count();
echo ':' . $collector->getErrors()[0]['msg'];
