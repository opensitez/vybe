<?php
// vybe-test: php/output_buffering/ob_clean_discards
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "will be discarded";
ob_clean();
echo "this survives";
ob_end_clean();
echo "done";
