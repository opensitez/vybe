<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_get_inner_iterator
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$inner = new ArrayIterator(["a", "b"]);
$it = new CachingIterator($inner);
echo $it->getInnerIterator() === $inner ? "INNER_ITERATOR_OK" : "FAIL";
