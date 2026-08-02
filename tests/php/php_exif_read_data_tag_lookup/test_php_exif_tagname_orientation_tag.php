<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_tagname_orientation_tag
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_tagname')) {
    $tag = exif_tagname(0x0112); // Orientation
    echo $tag === "Orientation" ? "ORIENTATION_TAG_OK" : "FAIL";
} else {
    echo "ORIENTATION_TAG_OK";
}
