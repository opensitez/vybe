<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_string_alphabetical_ascending
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert("charlie");
$heap->insert("alice");
$heap->insert("bob");
echo $heap->extract() === "alice" ? "ALICE_MIN_OK" : "FAIL";
