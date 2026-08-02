<?php
// vybe-test: php/string_builtins_extended/levenshtein_edit_distance
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo levenshtein("kitten", "sitting");
echo levenshtein("sunday", "saturday");
echo levenshtein("abc", "abc");
