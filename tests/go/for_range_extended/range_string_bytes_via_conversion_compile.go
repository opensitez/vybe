// vybe-test: go/for_range_extended/range_string_bytes_via_conversion_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i, b := range []byte("go") { _, _ = i, b } }
