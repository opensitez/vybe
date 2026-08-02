// vybe-test: go/strings_unicode_extra/string_map_lookup_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]string{"a": "b"}
_, _ = values["a"] }
