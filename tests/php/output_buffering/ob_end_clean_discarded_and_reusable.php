<?php
// vybe-test: php/output_buffering/ob_end_clean_discarded_and_reusable
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'discard';
ob_end_clean();
ob_start();
echo 'keep';
echo ob_get_clean();
