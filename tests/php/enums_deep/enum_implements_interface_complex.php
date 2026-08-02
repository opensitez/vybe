<?php
// vybe-test: php/enums_deep/enum_implements_interface_complex
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

interface HasDisplayName {
    public function displayName(): string;
}
interface HasIcon {
    public function icon(): string;
}
enum FileType: string implements HasDisplayName, HasIcon {
    case PDF   = 'pdf';
    case Word  = 'docx';
    case Excel = 'xlsx';
    public function displayName(): string {
        return match($this) {
            self::PDF   => 'PDF Document',
            self::Word  => 'Word Document',
            self::Excel => 'Excel Spreadsheet',
        };
    }
    public function icon(): string {
        return match($this) {
            self::PDF   => '📄',
            self::Word  => '📝',
            self::Excel => '📊',
        };
    }
}
echo FileType::PDF->displayName();
echo FileType::Word->value;
