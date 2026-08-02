// vybe-test: go/strings_unicode_extra/string_struct_field_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
type holder struct { text string }
func main() { _ = holder{text: "go"} }
