<?php
// vybe-test: php/clone_patterns/cloned_objects_are_equal_but_not_identical
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

class Tag { public string $name; public function __construct(string $n) { $this->name = $n; } }
$a = new Tag("php");
$b = clone $a;
echo ($a == $b ? 'equal' : 'not equal') . ',' . ($a === $b ? 'identical' : 'not identical');

__vybe_check(ob_get_clean(), "equal,not identical");
