<?php
// vybe-test: php/intl/collator_sort
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$words = ['Banana', 'apple', 'Cherry', 'date'];
$coll->sort($words);
echo implode(',', $words);
