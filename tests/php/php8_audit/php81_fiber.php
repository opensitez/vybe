<?php
// vybe-test: php/php8_audit/php81_fiber
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$f = new Fiber(function() { Fiber::suspend('hi'); }); echo $f->start();
