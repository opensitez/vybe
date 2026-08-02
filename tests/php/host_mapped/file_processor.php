<?php
// vybe-test: php/host_mapped/file_processor
// origin: languages/php/tests/php/test_host_mapped.rs
// vybe-test-mode: compile

$files = scandir('/tmp');
foreach ($files as $file) {
    if (is_file('/tmp/' . $file)) {
        $size = filesize('/tmp/' . $file);
        $ext = pathinfo($file);
        echo $file . ': ' . $size . ' bytes';
    }
}
