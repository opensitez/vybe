<?php
// vybe-test: php/array_extra_builtins/array_reduce_without_initial_and_empty_array_compiles
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

echo array_reduce([], fn($carry, $item) => $carry + $item);
