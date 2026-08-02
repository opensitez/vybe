<?php
// vybe-test: php/advanced_oop/fluent_validation_chain
// origin: languages/php/tests/php/test_advanced_oop.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

class Validator {
    private array $errors = [];
    private mixed $value;
    public function __construct(mixed $v) { $this->value = $v; }
    public function required(): static { if (empty($this->value)) $this->errors[] = 'required'; return $this; }
    public function minLength(int $n): static { if (strlen($this->value) < $n) $this->errors[] = "min:$n"; return $this; }
    public function isValid(): bool { return empty($this->errors); }
    public function errors(): array { return $this->errors; }
}
$v = new Validator('hi');
$v->required()->minLength(5);
echo $v->isValid() ? 'valid' : implode(',', $v->errors());

__vybe_check(ob_get_clean(), "min:5");
