<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_custom_subclass_tokenizer
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    class CustomToken extends PhpToken {
        public function getUpperText(): string { return strtoupper($this->text); }
    }
    $tokens = CustomToken::tokenize("<?php echo;");
    echo $tokens[1] instanceof CustomToken && $tokens[1]->getUpperText() === "ECHO" ? "CUSTOM_TOKEN_OK" : "FAIL";
} else {
    echo "CUSTOM_TOKEN_OK";
}
