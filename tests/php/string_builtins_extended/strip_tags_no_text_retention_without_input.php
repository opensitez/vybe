<?php
// vybe-test: php/string_builtins_extended/strip_tags_no_text_retention_without_input
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo strip_tags("<script>alert(1)</script>");
echo "|";
echo strip_tags("<b>safe</b><i>ok</i>", "<i>");
