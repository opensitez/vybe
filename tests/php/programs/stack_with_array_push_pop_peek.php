<?php
// vybe-test: php/programs/stack_with_array_push_pop_peek
// origin: languages/php/tests/php/test_programs.rs

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

class Stack {
    private array $data = [];
    public function push($v): void { $this->data[] = $v; }
    public function pop() { return array_pop($this->data); }
    public function peek() { return end($this->data); }
    public function isEmpty(): bool { return empty($this->data); }
    public function size(): int { return count($this->data); }
}
$s = new Stack();
$s->push(1); $s->push(2); $s->push(3);
echo $s->peek() . "\n";
echo $s->pop() . "\n";
echo $s->size() . "\n";

__vybe_check(ob_get_clean(), "3\n3\n2");
