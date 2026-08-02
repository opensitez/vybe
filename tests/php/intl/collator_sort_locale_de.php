<?php
// vybe-test: php/intl/collator_sort_locale_de
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('de_DE');
$words = ['Österreich', 'Angola', 'Zürich', 'Belgien'];
$coll->sort($words);
echo implode(',', $words);
