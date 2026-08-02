<?php
// vybe-test: php/trait_conflict_resolution/trait_abstract_method_fails_validation_on_missing_field
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
    abstract protected function required(): array;
    public function check(array $data): string {
        foreach ($this->required() as $field) {
            if (!isset($data[$field])) return "missing: $field";
        }
        return "ok";
    }
}
class UserForm {
    use Validatable;
    protected function required(): array { return ['name', 'age']; }
}
echo (new UserForm())->check(['name' => 'Bob']);

__vybe_check(ob_get_clean(), "missing: age");
