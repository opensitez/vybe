<?php
// vybe-test: php/classes/class_copy_by_reference_with_cloning
// origin: languages/php/tests/php/test_classes.rs

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

class Box { public string $label; public function __construct(string $label) { $this->label = $label; } }
$b1 = new Box('x');
$b2 = clone $b1;
$b2->label = 'y';
echo $b1->label, '|', $b2->label;

__vybe_check(ob_get_clean(), "x|y");
