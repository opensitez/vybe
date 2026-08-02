<?php
// vybe-test: php/types/gettype_check
// origin: languages/php/tests/php/test_types.rs
// vybe-test-mode: compile

echo gettype(42); echo gettype('hi'); echo gettype(null); echo gettype(true); echo gettype([]);
