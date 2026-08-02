<?php
// vybe-test: php/goto/goto_nested_function
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

function process(bool $skip): string {
    if ($skip) goto done;
    $result = 'processed';
    goto end;
    done:
    $result = 'skipped';
    end:
    return $result;
}
echo process(false);
echo process(true);
