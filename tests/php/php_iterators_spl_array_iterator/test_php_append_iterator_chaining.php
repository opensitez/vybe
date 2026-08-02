<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_append_iterator_chaining
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$app = new AppendIterator();
$app->append(new ArrayIterator([1, 2]));
$app->append(new ArrayIterator([3, 4]));
echo implode(",", iterator_to_array($app, false));
