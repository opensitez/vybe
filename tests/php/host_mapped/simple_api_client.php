<?php
// vybe-test: php/host_mapped/simple_api_client
// origin: languages/php/tests/php/test_host_mapped.rs
// vybe-test-mode: compile

$response = file_get_contents('https://api.example.com/data');
$data = json_decode($response);
echo json_encode($data);
