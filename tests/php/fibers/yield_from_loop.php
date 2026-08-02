<?php
// vybe-test: php/fibers/yield_from_loop
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function range_gen($start, $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
