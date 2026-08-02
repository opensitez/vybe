<?php
// vybe-test: php/references/global_reference_accumulate
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$sum = 0;
function accumulate(int $n) {
    global $sum;
    $sum += $n;
}
foreach ([10, 20, 30] as $v) { accumulate($v); }
echo $sum;
