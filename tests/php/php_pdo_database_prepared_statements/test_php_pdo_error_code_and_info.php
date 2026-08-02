<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_error_code_and_info
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
@$pdo->query("SELECT * FROM non_existent_table");
echo "ErrCode: " . $pdo->errorCode();
print_r($pdo->errorInfo());
