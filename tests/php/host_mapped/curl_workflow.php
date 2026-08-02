<?php
// vybe-test: php/host_mapped/curl_workflow
// origin: languages/php/tests/php/test_host_mapped.rs
// vybe-test-mode: compile

$ch = curl_init();
curl_setopt($ch, 'CURLOPT_URL', 'https://example.com');
$result = curl_exec($ch);
curl_close($ch);
