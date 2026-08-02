<?php
// vybe-test: php/programs/queue_with_array_enqueue_dequeue
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

class Queue {
    private array $data = [];
    public function enqueue($v): void { $this->data[] = $v; }
    public function dequeue() { return array_shift($this->data); }
    public function front() { return $this->data[0] ?? null; }
    public function size(): int { return count($this->data); }
}
$q = new Queue();
$q->enqueue('a'); $q->enqueue('b'); $q->enqueue('c');
echo $q->dequeue() . "\n";
echo $q->front() . "\n";
echo $q->size() . "\n";

__vybe_check(ob_get_clean(), "a\nb\n2");
