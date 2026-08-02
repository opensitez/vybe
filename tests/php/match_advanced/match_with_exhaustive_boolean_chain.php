<?php
// vybe-test: php/match_advanced/match_with_exhaustive_boolean_chain
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$v = match (true) {
    true === true => 'true-branch',
    false => 'false-branch',
};
echo $v;
