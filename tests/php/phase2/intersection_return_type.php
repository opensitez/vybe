<?php
// vybe-test: php/phase2/intersection_return_type
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

function bar(): Countable&Iterator { return null; }
