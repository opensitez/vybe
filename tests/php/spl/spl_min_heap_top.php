<?php
// vybe-test: php/spl/spl_min_heap_top
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
foreach ([30, 10, 50, 20] as $v) { $heap->insert($v); }
echo $heap->top();
