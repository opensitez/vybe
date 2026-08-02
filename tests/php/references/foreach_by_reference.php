<?php
// vybe-test: php/references/foreach_by_reference
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$arr = [1, 2, 3];
foreach ($arr as &$v) { $v *= 2; }
unset($v);
echo implode(',', $arr);
