<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_count_mode
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator([1, 2, 3, 4]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
echo count($it) === 4 ? "COUNT_CACHE_OK" : "FAIL";
