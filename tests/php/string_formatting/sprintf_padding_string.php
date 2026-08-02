<?php
// vybe-test: php/string_formatting/sprintf_padding_string
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%10s", "hello");  // "     hello"
echo sprintf("%-10s|", "hi");   // "hi        |"
echo sprintf("%'#10s", "ok");   // "########ok"
