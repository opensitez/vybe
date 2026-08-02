// vybe-test: go/strings_unicode_extra/string_range_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
func main() { for i, r := range "go" { _, _ = i, r } }
