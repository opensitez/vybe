<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_serialize_unserialize
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["x", "y"]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
$s = serialize($it);
$restored = unserialize($s);
echo count($restored) === 2 ? "SERIALIZE_CACHING_OK" : "FAIL";
