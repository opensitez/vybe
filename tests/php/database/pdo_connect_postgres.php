<?php
// vybe-test: php/database/pdo_connect_postgres
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('postgresql://localhost/mydb');
