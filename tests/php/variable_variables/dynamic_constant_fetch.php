<?php
// vybe-test: php/variable_variables/dynamic_constant_fetch
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Direction {
    const NORTH = 'N';
    const SOUTH = 'S';
    const EAST  = 'E';
    const WEST  = 'W';
}
foreach (['NORTH', 'SOUTH', 'EAST', 'WEST'] as $dir) {
    echo Direction::$$dir ?? constant("Direction::$dir");
}
