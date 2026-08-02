<?php
// vybe-test: php/cross_lang/hashset
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$set = new HashSet();
$set->add('apple');
$set->add('banana');
$set->add('apple'); // duplicate
echo $set->contains('banana');
$set->remove('banana');
