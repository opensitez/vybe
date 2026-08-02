<?php
// vybe-test: php/generators_advanced/generator_return_type_hint
// origin: languages/php/tests/php/test_generators_advanced.rs
// vybe-test-mode: compile

function counter(int $start, int $end): Generator {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
$g = counter(1, 3);
foreach ($g as $v) {
    echo $v;
}
