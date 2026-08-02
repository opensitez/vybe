// vybe-test: go/strings_unicode_extra/string_type_alias_map_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
type label string
func main() { values := map[label]int{"go": 1}
_ = values }
