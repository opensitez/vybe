<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_flags_setter_getter
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["test"]);
$it = new CachingIterator($arr);
$it->setFlags(CachingIterator::CALL_TOSTRING);
echo ($it->getFlags() & CachingIterator::CALL_TOSTRING) ? "FLAGS_OK" : "FAIL";
