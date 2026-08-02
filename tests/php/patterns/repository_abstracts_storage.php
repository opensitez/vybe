<?php
// vybe-test: php/patterns/repository_abstracts_storage
// origin: languages/php/tests/php/test_patterns.rs

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

class User {
    public function __construct(public int $id, public string $name) {}
}
class UserRepository {
    private $store = [];
    public function save(User $u): void { $this->store[$u->id] = $u; }
    public function find(int $id): ?User { return $this->store[$id] ?? null; }
    public function findAll(): array { return array_values($this->store); }
}
$repo = new UserRepository();
$repo->save(new User(1, 'Alice'));
$repo->save(new User(2, 'Bob'));
echo $repo->find(1)->name;
echo count($repo->findAll());

__vybe_check(ob_get_clean(), "Alice2");
