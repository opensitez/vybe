<?php
// vybe-test: php/spl_extra/spl_min_heap_count_after_extractions
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$h = new SplMinHeap();
$h->insert(5); $h->insert(3); $h->insert(8); $h->insert(1);
echo $h->count();
$h->extract(); $h->extract();
echo $h->count();
