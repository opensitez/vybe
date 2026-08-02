<?php
// vybe-test: php/error_handling_deep/parse_error_via_eval
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

try {
    eval('$x = ;'); // parse error
} catch (\ParseError $e) {
    echo 'parse error caught';
}
