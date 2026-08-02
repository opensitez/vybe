<?php
// vybe-test: php/database/mysqli_close
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$conn = mysqli_connect('sqlite:test.db');
mysqli_close($conn);
