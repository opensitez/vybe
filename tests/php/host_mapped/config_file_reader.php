<?php
// vybe-test: php/host_mapped/config_file_reader
// origin: languages/php/tests/php/test_host_mapped.rs
// vybe-test-mode: compile

$content = file_get_contents('config.json');
$config = json_decode($content);
$dbDsn = 'sqlite:' . getcwd() . '/app.db';
$pdo = new PDO($dbDsn);
