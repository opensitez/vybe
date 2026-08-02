<?php
// vybe-test: php/spl_extra/spl_max_heap_extract_order
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$h = new SplMaxHeap();
foreach ([9, 3, 7, 1, 5] as $v) { $h->insert($v); }
$out = [];
while (!$h->isEmpty()) { $out[] = $h->extract(); }
echo implode(',', $out); // 9,7,5,3,1
