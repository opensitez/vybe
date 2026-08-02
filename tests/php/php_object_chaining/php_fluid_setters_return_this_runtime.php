<?php
// vybe-test: php/php_object_chaining/php_fluid_setters_return_this_runtime
// origin: languages/php/tests/php/test_php_object_chaining.rs

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

class Profile {
    public string $name = '';
    public string $role = '';
    public function set_name(string $name): self { $this->name = $name; return $this; }
    public function set_role(string $role): self { $this->role = $role; return $this; }
    public function label(): string { return $this->name . ':' . $this->role; }
}
$p = (new Profile())->set_name('alice')->set_role('admin');
echo $p->label();

__vybe_check(ob_get_clean(), "alice:admin");
