<?php
// vybe-test: php/intl/collator_strength
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$coll->setStrength(Collator::PRIMARY);
echo $coll->compare('Cafe', 'café') === 0 ? 'equal (primary)' : 'different';
