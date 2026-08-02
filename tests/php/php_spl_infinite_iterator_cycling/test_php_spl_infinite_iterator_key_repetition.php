<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_key_repetition
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["k1" => "v1", "k2" => "v2"]);
$inf = new InfiniteIterator($arr);
$keys = [];
$i = 0;
foreach ($inf as $k => $v) {
    $keys[] = $k;
    if (++$i >= 4) break;
}
echo implode(",", $keys) === "k1,k2,k1,k2" ? "KEYS_REPEAT_OK" : "FAIL";
