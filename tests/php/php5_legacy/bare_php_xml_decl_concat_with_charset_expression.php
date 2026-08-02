<?php
// vybe-test: php/php5_legacy/bare_php_xml_decl_concat_with_charset_expression
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

$xml = '<?xml version="1.0" encoding="'.$charset.'"?>'."\n".$xml;
