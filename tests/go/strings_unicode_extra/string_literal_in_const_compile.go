// vybe-test: go/strings_unicode_extra/string_literal_in_const_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
const a = "go"
const b = a + "lang"
func main() { _ = b }
