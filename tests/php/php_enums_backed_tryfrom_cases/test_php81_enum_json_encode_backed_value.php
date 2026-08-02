<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_json_encode_backed_value
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Status: string { case Active = "active"; }
echo json_encode(Status::Active); // Enums serialize to backed value or name
