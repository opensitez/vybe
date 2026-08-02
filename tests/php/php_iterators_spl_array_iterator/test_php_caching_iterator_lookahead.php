<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_caching_iterator_lookahead
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$cit = new CachingIterator(new ArrayIterator([1, 2, 3]), CachingIterator::FULL_CACHE);
foreach ($cit as $val) {
    if (!$cit->hasNext()) {
        echo "LAST:$val";
    }
}
