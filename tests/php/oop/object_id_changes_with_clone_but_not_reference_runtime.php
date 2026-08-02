<?php
// vybe-test: php/oop/object_id_changes_with_clone_but_not_reference_runtime
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

class Holder {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
}
$a = new Holder('left');
$b = $a;
$c = clone $a;
$b->name = 'right';
echo spl_object_id($a) === spl_object_id($b) ? 'same' : 'diff';
echo '|';
echo spl_object_id($a) === spl_object_id($c) ? 'same' : 'diff';

__vybe_check(ob_get_clean(), "same|diff");
