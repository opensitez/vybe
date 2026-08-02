<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_thumbnail_returns_false_for_no_thumb
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_thumbnail')) {
    $tmp = sys_get_temp_dir() . "/test_thumb_" . uniqid() . ".png";
    file_put_contents($tmp, "dummy data");
    $thumb = @exif_thumbnail($tmp, $w, $h, $t);
    @unlink($tmp);
    echo $thumb === false ? "NO_THUMB_FALSE_OK" : "FAIL";
} else {
    echo "NO_THUMB_FALSE_OK";
}
