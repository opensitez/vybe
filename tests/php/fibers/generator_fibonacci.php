<?php
// vybe-test: php/fibers/generator_fibonacci
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function fibonacci() {
    $a = 0;
    $b = 1;
    while (true) {
        yield $a;
        $tmp = $a;
        $a = $b;
        $b = $tmp + $b;
    }
}
