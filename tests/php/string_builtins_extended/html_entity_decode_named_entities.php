<?php
// vybe-test: php/string_builtins_extended/html_entity_decode_named_entities
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo html_entity_decode("&lt;b&gt;bold&lt;/b&gt; &amp; &copy; &trade;");
