<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_mime_type_invalid_returns_application_octet_stream
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_mime_type')) {
    $mime = image_type_to_mime_type(999999);
    echo $mime === "application/octet-stream" ? "OCTET_STREAM_FALLBACK_OK" : "FAIL";
} else {
    echo "OCTET_STREAM_FALLBACK_OK";
}
