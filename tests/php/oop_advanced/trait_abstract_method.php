<?php
// vybe-test: php/oop_advanced/trait_abstract_method
// origin: languages/php/tests/php/test_oop_advanced.rs

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

trait Validatable {
    abstract protected function validate(): bool;
    public function isValid(): string {
        return $this->validate() ? "valid" : "invalid";
    }
}
class Email {
    use Validatable;
    public function __construct(private string $value) {}
    protected function validate(): bool {
        return str_contains($this->value, "@");
    }
}
$e1 = new Email("user@example.com");
$e2 = new Email("invalid");
echo $e1->isValid(), "\n";
echo $e2->isValid(), "\n";

__vybe_check(ob_get_clean(), "valid\ninvalid");
