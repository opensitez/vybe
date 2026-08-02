<?php
// vybe-test: php/spl_extra/spl_min_heap_extract_order
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$h = new SplMinHeap();
foreach ([9, 3, 7, 1, 5] as $v) { $h->insert($v); }
$out = [];
while (!$h->isEmpty()) { $out[] = $h->extract(); }
echo implode(',', $out); // 1,3,5,7,9
