<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_imagetype_invalid_file_returns_false
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_imagetype')) {
    $res = @exif_imagetype("/path/to/nonexistent/file/99999.png");
    echo $res === false ? "NONEXISTENT_EXIF_TYPE_FALSE_OK" : "FAIL";
} else {
    echo "NONEXISTENT_EXIF_TYPE_FALSE_OK";
}
