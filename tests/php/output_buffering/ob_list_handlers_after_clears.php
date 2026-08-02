<?php
// vybe-test: php/output_buffering/ob_list_handlers_after_clears
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start(function(string $b): string { return '[' . $b . ']'; });
$before = count(ob_list_handlers());
ob_clean();
$after = count(ob_list_handlers());
ob_end_clean();
ob_end_clean();
echo $before . '|' . $after;
