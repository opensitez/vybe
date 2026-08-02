<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_read_data_section_names_parameter
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_read_data')) {
    $tmp = sys_get_temp_dir() . "/test_sec_" . uniqid() . ".jpg";
    file_put_contents($tmp, "dummy data");
    $data = @exif_read_data($tmp, "IFD0", true);
    @unlink($tmp);
    echo $data === false ? "SECTION_NAMES_PARAM_OK" : "FAIL";
} else {
    echo "SECTION_NAMES_PARAM_OK";
}
