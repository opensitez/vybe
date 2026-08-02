<?php
// vybe-test: php/string_extra_builtins/html_entity_decode_chars_from_entities
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$encoded = "&lt;p&gt;Hello &amp; World&lt;/p&gt;";
$decoded = html_entity_decode($encoded);
echo strpos($decoded, "<p>") !== false ? "has-tag" : "no-tag";
echo strpos($decoded, "&") !== false ? "has-amp" : "no-amp";
