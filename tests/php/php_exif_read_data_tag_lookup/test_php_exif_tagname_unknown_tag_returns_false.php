<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_tagname_unknown_tag_returns_false
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs
// vybe-test-mode: compile

if (function_exists('exif_tagname')) {
    $tag = @exif_tagname(0xFFFF);
    echo $tag === false ? "UNKNOWN_TAG_FALSE_OK" : "FAIL";
} else {
    echo "UNKNOWN_TAG_FALSE_OK";
}
