<?php
// vybe-test: php/string_builtins_extended/similar_text_comparison
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$common = similar_text("World", "Word");
echo $common;
similar_text("Hello", "Hello", $pct);
echo ($pct == 100.0) ? "full" : "partial";
