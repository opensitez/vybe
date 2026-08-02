<?php
// vybe-test: php/cross_lang/stringbuilder_replace
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$sb = new StringBuilder('Hello World');
$sb->replace('World', 'PHP');
echo $sb->toString();
$sb->clear();
