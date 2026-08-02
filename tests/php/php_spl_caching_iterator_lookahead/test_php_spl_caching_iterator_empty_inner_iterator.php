<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_empty_inner_iterator
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator([]);
$it = new CachingIterator($arr);
echo !$it->hasNext() ? "EMPTY_HAS_NEXT_FALSE" : "FAIL";
