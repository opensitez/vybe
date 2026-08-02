<?php
// vybe-test: php/database/pdo_connect_sqlite
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:test.db');
