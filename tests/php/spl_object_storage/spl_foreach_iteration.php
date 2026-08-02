<?php
// vybe-test: php/spl_object_storage/spl_foreach_iteration
// origin: languages/php/tests/php/test_spl_object_storage.rs

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

class Tag { public function __construct(public string $name) {} }
$s = new SplObjectStorage;
$s->attach(new Tag('a'), 1);
$s->attach(new Tag('b'), 2);
$s->attach(new Tag('c'), 3);
$sum = 0;
foreach ($s as $obj) { $sum += $s->getInfo(); }
echo $sum;

__vybe_check(ob_get_clean(), "6");
