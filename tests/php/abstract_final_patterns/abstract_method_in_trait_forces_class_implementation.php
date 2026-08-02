<?php
// vybe-test: php/abstract_final_patterns/abstract_method_in_trait_forces_class_implementation
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

trait Validator {
    abstract protected function rules(): array;
    public function validate(array $data): bool {
        foreach ($this->rules() as $rule) if (!isset($data[$rule])) return false;
        return true;
    }
}
class Form {
    use Validator;
    protected function rules(): array { return ['email', 'password']; }
}
$f = new Form();
echo $f->validate(['email' => 'a@b.com', 'password' => 'x']) ? 'valid' : 'invalid', "\n";

__vybe_check(ob_get_clean(), "valid");
