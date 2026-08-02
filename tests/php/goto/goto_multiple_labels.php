<?php
// vybe-test: php/goto/goto_multiple_labels
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

$step = 2;
goto {"step$step"};
step1: echo "1"; goto end;
step2: echo "2"; goto end;
step3: echo "3"; goto end;
end:
