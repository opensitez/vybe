<?php
// vybe-test: php/goto/goto_skip_block
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

$x = 1;
if ($x > 0) { goto positive; }
echo "non-positive";
goto done;
positive:
echo "positive";
done:
