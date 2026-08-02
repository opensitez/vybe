<?php
// vybe-test: php/intl/collator_basic_compare
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
echo $coll->compare('apple', 'Banana') < 0 ? 'apple < Banana' : 'apple >= Banana';
