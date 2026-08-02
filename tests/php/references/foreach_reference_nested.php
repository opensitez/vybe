<?php
// vybe-test: php/references/foreach_reference_nested
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$matrix = [[1, 2], [3, 4]];
foreach ($matrix as &$row) {
    foreach ($row as &$cell) { $cell += 10; }
    unset($cell);
}
unset($row);
echo $matrix[0][0] . ',' . $matrix[1][1];
