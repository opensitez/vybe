<?php
// vybe-test: php/declare/declare_ticks
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

$tick_count = 0;
register_tick_function(function() use (&$tick_count) { $tick_count++; });
declare(ticks=1) {
    for ($i = 0; $i < 5; $i++) { $x = $i * 2; }
}
echo $tick_count > 0 ? 'ticked' : 'no ticks';
