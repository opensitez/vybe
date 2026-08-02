<?php
// vybe-test: php/oop_patterns/interface_typehint_polymorphism
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

interface Formatter {
    public function format(mixed $value): string;
}
class NumberFormatter implements Formatter {
    public function format(mixed $value): string { return number_format((float)$value, 2); }
}
class UpperFormatter implements Formatter {
    public function format(mixed $value): string { return strtoupper((string)$value); }
}
function render(Formatter $fmt, mixed $val): void {
    echo $fmt->format($val);
}
render(new NumberFormatter(), 1234.5);
render(new UpperFormatter(), 'hello');
