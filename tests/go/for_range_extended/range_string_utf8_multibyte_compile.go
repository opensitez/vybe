// vybe-test: go/for_range_extended/range_string_utf8_multibyte_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i, r := range "é" { _, _ = i, r } }
