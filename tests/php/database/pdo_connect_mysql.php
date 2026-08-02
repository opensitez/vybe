<?php
// vybe-test: php/database/pdo_connect_mysql
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('mysql:host=localhost;dbname=mydb');
