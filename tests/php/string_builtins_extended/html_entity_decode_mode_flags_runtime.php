<?php
// vybe-test: php/string_builtins_extended/html_entity_decode_mode_flags_runtime
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo html_entity_decode('&quot;A&amp;B&quot;', ENT_QUOTES);
echo "|";
echo html_entity_decode('&#x41;&#x42;', ENT_NOQUOTES);
