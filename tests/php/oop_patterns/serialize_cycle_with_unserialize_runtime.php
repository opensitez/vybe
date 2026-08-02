<?php
// vybe-test: php/oop_patterns/serialize_cycle_with_unserialize_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Box {
    public function __construct(public int $n, public string $label) {}
    public function __serialize(): array {
        return ['n' => $this->n, 'label' => $this->label];
    }
    public function __unserialize(array $data): void {
        $this->n = $data['n'];
        $this->label = $data['label'] . '!'; 
    }
}
$box = new Box(7, 'ok');
$text = serialize($box);
$copy = unserialize($text);
echo $copy->n . '|' . $copy->label;

__vybe_check(ob_get_clean(), "7|ok!");
