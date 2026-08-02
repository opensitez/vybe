// vybe-test: go/strings_unicode_extra/string_concat_with_named_type_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
type label string
func main() { var left label = "go"
_ = string(left) + "lang" }
