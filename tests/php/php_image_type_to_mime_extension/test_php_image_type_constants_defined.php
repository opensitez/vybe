<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_constants_defined
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

$hasConsts = defined('IMAGETYPE_GIF') && defined('IMAGETYPE_JPEG') && defined('IMAGETYPE_PNG') && defined('IMAGETYPE_WEBP');
echo $hasConsts ? "IMAGETYPE_CONSTANTS_DEFINED" : "FAIL";
