<?php
// vybe-test: php/php81_features/fiber_passes_value_on_suspend
// origin: languages/php/tests/php/test_php81_features.rs

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

$f = new Fiber(function(): string {
    $v = Fiber::suspend('mid');
    return 'end:' . $v;
});
$mid = $f->start();
echo $mid . ',';
$f->resume('resumed');
echo $f->getReturn();

__vybe_check(ob_get_clean(), "mid,end:resumed");
