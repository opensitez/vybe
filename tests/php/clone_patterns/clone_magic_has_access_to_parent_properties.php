<?php
// vybe-test: php/clone_patterns/clone_magic_has_access_to_parent_properties
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class BaseModel { protected string $createdAt = '2024-01-01'; }
class User extends BaseModel {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
    public function __clone() { $this->createdAt = '2024-06-01'; }
    public function getCreated(): string { return $this->createdAt; }
}
$u = new User("Alice");
$v = clone $u;
echo $u->getCreated() . ',' . $v->getCreated();

__vybe_check(ob_get_clean(), "2024-01-01,2024-06-01");
