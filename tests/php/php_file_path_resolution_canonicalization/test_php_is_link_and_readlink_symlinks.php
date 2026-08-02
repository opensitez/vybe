<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_is_link_and_readlink_symlinks
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$target = tempnam(sys_get_temp_dir(), "vybe_target_");
$link = sys_get_temp_dir() . "/vybe_symlink_" . time();
@symlink($target, $link);
if (is_link($link)) {
    echo "SYMLINK_CREATED: " . readlink($link);
    unlink($link);
}
unlink($target);
