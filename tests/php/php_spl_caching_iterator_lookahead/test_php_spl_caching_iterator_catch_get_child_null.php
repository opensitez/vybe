<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_catch_get_child_null
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator([1]);
$it = new CachingIterator($arr);
echo $it->getChildren() === null ? "NO_CHILDREN_OK" : "FAIL";
