<?php
// vybe-test: php/magic_methods/magic_get_returning_object_with_get
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Nested {
    private array $children = [];
    public function __construct(private string $name) {}
    public function addChild(string $key, Nested $child): void {
        $this->children[$key] = $child;
    }
    public function __get($key): ?Nested {
        return $this->children[$key] ?? null;
    }
    public function getName(): string { return $this->name; }
}
$root = new Nested("root");
$root->addChild("child", new Nested("child_node"));
echo $root->child->getName();

__vybe_check(ob_get_clean(), "child_node");
