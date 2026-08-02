<?php
// vybe-test: php/functional_style/generator_lazy_range
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function lazyRange(int $start, int $end): Generator {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
$gen = lazyRange(1, 5);
foreach ($gen as $n) {
    echo $n;
}
