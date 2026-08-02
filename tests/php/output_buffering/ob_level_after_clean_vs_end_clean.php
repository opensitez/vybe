<?php
// vybe-test: php/output_buffering/ob_level_after_clean_vs_end_clean
// origin: languages/php/tests/php/test_output_buffering.rs

$base = ob_get_level();
ob_start();
$a = ob_get_level() === $base + 1 ? 'nested' : 'wrong';
ob_clean();
$b = ob_get_level() === $base + 1 ? 'still' : 'gone';
ob_end_clean();
$c = ob_get_level() === $base ? 'done' : 'open';
echo $a . '|' . $b . '|' . $c;
