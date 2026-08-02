<?php
// vybe-test: php/cross_lang/await_promise
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

// await() uses same opcode as JS await — can await JS promises
$result = await(fetch('https://api.example.com/data'));
echo $result;
