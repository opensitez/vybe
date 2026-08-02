<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_offset_get_cache_key
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["key1" => "val1", "key2" => "val2"]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
echo $it["key1"] === "val1" ? "OFFSET_GET_CACHE_OK" : "FAIL";
