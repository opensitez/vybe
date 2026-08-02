<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_unexpected_value_exception_parsing
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

function parseFormat(string $format) {
    if ($format !== "json" && $format !== "xml") {
        throw new UnexpectedValueException("Expected json or xml, got $format");
    }
}

try {
    parseFormat("yaml");
} catch (UnexpectedValueException $e) {
    echo $e->getMessage();
}
