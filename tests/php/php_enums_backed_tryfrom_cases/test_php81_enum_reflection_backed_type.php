<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_reflection_backed_type
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Code: int { case OK = 200; }
$re = new ReflectionEnum(Code::class);
$type = $re->getBackingType();
echo $type->getName();
