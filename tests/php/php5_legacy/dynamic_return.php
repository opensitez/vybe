<?php
// vybe-test: php/php5_legacy/dynamic_return
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function parse($input) {
    if (is_numeric($input)) return intval($input);
    if ($input === 'true') return true;
    if ($input === 'null') return null;
    return $input;
}
echo parse('42');
echo parse('hello');
