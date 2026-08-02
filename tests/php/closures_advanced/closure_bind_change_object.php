<?php
// vybe-test: php/closures_advanced/closure_bind_change_object
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Counter { private int $count = 0; }
$inc = Closure::bind(function() { $this->count++; return $this->count; }, new Counter(), Counter::class);
echo $inc();
echo $inc();
