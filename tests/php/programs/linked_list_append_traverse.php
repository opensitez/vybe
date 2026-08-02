<?php
// vybe-test: php/programs/linked_list_append_traverse
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

class Node {
    public $next = null;
    public function __construct(public $value) {}
}
class LinkedList {
    private $head = null;
    public function append($v): void {
        $node = new Node($v);
        if ($this->head === null) { $this->head = $node; return; }
        $cur = $this->head;
        while ($cur->next !== null) $cur = $cur->next;
        $cur->next = $node;
    }
    public function toArray(): array {
        $res = []; $cur = $this->head;
        while ($cur !== null) { $res[] = $cur->value; $cur = $cur->next; }
        return $res;
    }
}
$l = new LinkedList();
$l->append(10); $l->append(20); $l->append(30);
echo implode('->', $l->toArray()) . "\n";

__vybe_check(ob_get_clean(), "10->20->30");
