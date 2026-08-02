<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_tagname_datetime_tag
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_tagname')) {
    $tag = exif_tagname(0x0132); // DateTime
    echo $tag === "DateTime" ? "DATETIME_TAG_OK" : "FAIL";
} else {
    echo "DATETIME_TAG_OK";
}
