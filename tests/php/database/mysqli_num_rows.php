<?php
// vybe-test: php/database/mysqli_num_rows
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$conn = mysqli_connect('sqlite:test.db');
$result = mysqli_query($conn, 'SELECT * FROM users');
echo mysqli_num_rows($result);
