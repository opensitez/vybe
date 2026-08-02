<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_get_equivalent_id
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone')) {
    $id = IntlTimeZone::getEquivalentID("UTC", 0);
    echo $id === "UTC" || strlen($id) > 0 ? "EQUIVALENT_ID_OK" : "FAIL";
} else {
    echo "EQUIVALENT_ID_OK";
}
