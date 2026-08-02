<?php
// vybe-test: php/string_builtins_extended/strip_tags_allow_list
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$html = "<h1>Title</h1><p>Body <b>text</b></p><script>alert(1)</script>";
echo strip_tags($html, "<h1><p>");
