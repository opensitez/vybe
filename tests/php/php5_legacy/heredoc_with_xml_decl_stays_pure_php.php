<?php
// vybe-test: php/php5_legacy/heredoc_with_xml_decl_stays_pure_php
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class IXR_Request {
    public $xml = '';

    public function __construct() {
        $this->xml = <<<EOD
<?xml version="1.0"?>
<methodCall>
<params>
EOD;
        $this->xml .= '<param><value>';
    }
}
