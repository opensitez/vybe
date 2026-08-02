<?php
// vybe-test: php/enums_deep/backed_enum_try_from_invalid
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Code: int { case OK = 200; case NotFound = 404; case Error = 500; }
$found  = Code::tryFrom(200);
$missing = Code::tryFrom(999);
echo ($found !== null ? $found->name : 'null') . ':';
echo ($missing !== null ? $missing->name : 'null');
