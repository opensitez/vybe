<?php
// vybe-test: php/database/mysqli_surface_compile
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$dbh = mysqli_init();
mysqli_select_db($dbh, 'app');
mysqli_set_charset($dbh, 'utf8mb4');
mysqli_ping($dbh);
mysqli_errno($dbh);
mysqli_affected_rows($dbh);
mysqli_insert_id($dbh);
mysqli_num_fields($dbh);
mysqli_fetch_field($dbh);
mysqli_free_result($dbh);
mysqli_more_results($dbh);
mysqli_next_result($dbh);
mysqli_close($dbh);
mysqli_real_escape_string($dbh, 'hello');
mysqli_character_set_name($dbh);
mysqli_get_client_info();
mysqli_get_server_info($dbh);
