<?php
// vybe-test: php/declare/declare_ticks_block_form
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(ticks=1) {
    $sum = 0;
    for ($i = 1; $i <= 10; $i++) { $sum += $i; }
    echo $sum;
}
