<?php
// vybe-test: php/goto/goto_in_loop_exit
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

$sum = 0;
for ($i = 1; $i <= 10; $i++) {
    if ($i > 5) goto done;
    $sum += $i;
}
done:
echo $sum;
