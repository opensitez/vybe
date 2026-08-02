<?php
// vybe-test: php/oop/oop_magic_clone_independence_runtime
// origin: languages/php/tests/php/test_oop.rs

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
    public array $tags;
    public function __construct(array $tags) { $this->tags = $tags; }
    public function __clone(): void {
        $this->tags[] = 'cloned';
    }
}
$a = new Box(['a']);
$b = clone $a;
$a->tags[] = 'source';
echo implode(',', $a->tags);
echo '|';
echo implode(',', $b->tags);

__vybe_check(ob_get_clean(), "a,source|a,cloned");
