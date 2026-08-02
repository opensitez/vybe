// vybe-test: go/strings_unicode_extra/string_concat_in_return_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
func label() string { return "go" + "lang" }
func main() { _ = label() }
