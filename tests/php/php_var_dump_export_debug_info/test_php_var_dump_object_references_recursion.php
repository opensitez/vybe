<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_var_dump_object_references_recursion
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

$node1 = new stdClass();
$node2 = new stdClass();
$node1->next = $node2;
$node2->prev = $node1; // Circular reference

ob_start();
var_dump($node1);
$dump = ob_get_clean();
echo str_contains($dump, "*RECURSION*") ? "RECURSION_DETECTED" : "DUMP_OK";
