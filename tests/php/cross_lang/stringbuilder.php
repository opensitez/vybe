<?php
// vybe-test: php/cross_lang/stringbuilder
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$sb = new StringBuilder('Hello');
$sb->append(' World');
$sb->appendLine('!');
$sb->insert(5, ',');
echo $sb->toString();
