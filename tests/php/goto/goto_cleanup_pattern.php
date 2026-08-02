<?php
// vybe-test: php/goto/goto_cleanup_pattern
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

$cleanup_needed = false;
$result = 'pending';
$data = [1, 2, -1, 3];
foreach ($data as $v) {
    if ($v < 0) {
        $cleanup_needed = true;
        goto cleanup;
    }
    $result = "ok:$v";
}
goto done;
cleanup:
$result = "cleaned up";
done:
echo $result;
