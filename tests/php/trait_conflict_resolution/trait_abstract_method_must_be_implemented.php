<?php
// vybe-test: php/trait_conflict_resolution/trait_abstract_method_must_be_implemented
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs

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
    abstract protected function rules(): array;
    public function validate(array $data): bool {
        foreach ($this->rules() as $field) {
            if (empty($data[$field])) return false;
        }
        return true;
    }
}
class Form {
    use Validatable;
    protected function rules(): array { return ['name', 'email']; }
}
$f = new Form();
echo $f->validate(['name' => 'Alice', 'email' => 'a@b.com']) ? 'valid' : 'invalid';

__vybe_check(ob_get_clean(), "valid");
