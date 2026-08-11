<?php
// vybe-test: php/php_constants/define_with_trailing_comment
// origin: crates/vybe_compiler/src/bundle.rs (php_bundle_parses_define_with_trailing_inline_comment)
// vybe-test-mode: compile

define('ABSPATH', __DIR__ . '/');
define('WP_CONTENT_DIR', ABSPATH . 'wp-content'); // trailing comment
echo WP_CONTENT_DIR;
