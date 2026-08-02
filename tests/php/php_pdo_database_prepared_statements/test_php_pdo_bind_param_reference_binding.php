<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_bind_param_reference_binding
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE data (val INT)");
$stmt = $pdo->prepare("INSERT INTO data VALUES (?)");

$val = 0;
$stmt->bindParam(1, $val);
$val = 10; $stmt->execute();
$val = 20; $stmt->execute();

$stmt2 = $pdo->query("SELECT SUM(val) FROM data");
echo $stmt2->fetchColumn();
