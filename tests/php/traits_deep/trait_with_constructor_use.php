<?php
// vybe-test: php/traits_deep/trait_with_constructor_use
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Validatable {
    private array $errors = [];
    abstract protected function validate(): void;
    public function isValid(): bool { $this->validate(); return empty($this->errors); }
    protected function addError(string $msg): void { $this->errors[] = $msg; }
    public function getErrors(): array { return $this->errors; }
}
class Email {
    use Validatable;
    public function __construct(private string $addr) {}
    protected function validate(): void {
        if (!str_contains($this->addr, '@')) {
            $this->addError("Invalid email: {$this->addr}");
        }
    }
}
$e = new Email('notvalid');
echo $e->isValid() ? 'valid' : 'invalid';
echo ':' . count($e->getErrors());
