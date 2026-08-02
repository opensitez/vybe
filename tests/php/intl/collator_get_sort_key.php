<?php
// vybe-test: php/intl/collator_get_sort_key
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$key1 = $coll->getSortKey('apple');
$key2 = $coll->getSortKey('banana');
echo ($key1 < $key2) ? 'apple < banana' : 'apple >= banana';
