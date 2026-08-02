<?php
// vybe-test: php/goto/goto_conditional
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

$flag = true;
$result = '';
if ($flag) goto found;
$result = 'not found';
goto done;
found:
$result = 'found';
done:
echo $result;
