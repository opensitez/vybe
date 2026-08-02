<?php
// vybe-test: php/late_static_binding/lsb_fluent_builder_with_static_return
// origin: languages/php/tests/php/test_late_static_binding.rs

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

class Query {
    protected array $conditions = [];
    public function where(string $c): static { $this->conditions[] = $c; return $this; }
    public function build(): string { return implode(' AND ', $this->conditions); }
}
class UserQuery extends Query {}
echo (new UserQuery)->where('age>18')->where('active=1')->build();

__vybe_check(ob_get_clean(), "age>18 AND active=1");
