<?php
// vybe-test: php/intl/sort_names_locale_aware
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Collator')) { echo 'skipped'; return; }
$names = ['Müller', 'Maier', 'Möller', 'Meyer'];
$coll = new Collator('de_DE');
$coll->sort($names);
echo implode(',', $names);
