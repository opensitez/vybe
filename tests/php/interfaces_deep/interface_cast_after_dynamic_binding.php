<?php
// vybe-test: php/interfaces_deep/interface_cast_after_dynamic_binding
// origin: languages/php/tests/php/test_interfaces_deep.rs

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

interface IAnimal { public function kind(): string; }
interface IPet extends IAnimal { public function name(): string; }

class Cat implements IPet {
    public function __construct(private string $nameValue) {}
    public function kind(): string { return 'cat'; }
    public function name(): string { return $this->nameValue; }
}

$animal = new Cat('Misty');
echo ($animal instanceof IAnimal ? 'a' : 'x') . '-' . ($animal instanceof IPet ? 'p' : 'y');
echo '-' . $animal->kind() . ':' . $animal->name();

__vybe_check(ob_get_clean(), "a-p-cat:Misty");
