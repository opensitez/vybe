<?php
// vybe-test: php/goto/goto_forward_jump
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

echo "start";
goto step3;
echo "step2"; // skipped
step3:
echo "step3";
