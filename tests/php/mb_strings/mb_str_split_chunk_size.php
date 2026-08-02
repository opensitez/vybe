<?php
// vybe-test: php/mb_strings/mb_str_split_chunk_size
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$chunks = mb_str_split("Hello World", 3);
echo implode('|', $chunks);
