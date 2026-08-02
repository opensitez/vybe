<?php
// vybe-test: php/output_buffering/ob_nested_three_levels
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
    echo "L1";
    ob_start();
        echo "L2";
        ob_start();
            echo "L3";
        $l3 = ob_get_clean();
    $l2 = ob_get_clean();
$l1 = ob_get_clean();
echo "$l1-$l2-$l3";
