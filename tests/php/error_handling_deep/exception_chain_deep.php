<?php
// vybe-test: php/error_handling_deep/exception_chain_deep
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

function buildChain(int $depth, ?\Throwable $prev = null): \Throwable {
    if ($depth === 0) return new \RuntimeException("root", 0, $prev);
    return buildChain($depth - 1, new \RuntimeException("level $depth", 0, $prev));
}
$e = buildChain(3);
$count = 0;
while ($e !== null) { $count++; $e = $e->getPrevious(); }
echo $count;
