<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_foreach_iteration
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
$heap->insert(5);
$heap->insert(15);
$res = [];
foreach ($heap as $val) {
    $res[] = $val;
}
echo implode(",", $res) === "15,5" ? "FOREACH_HEAP_OK" : "FAIL";
