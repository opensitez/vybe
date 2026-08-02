// vybe-test: go/blank_identifier_extended/blank_discard_range_over_int_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { for range 2 { _ = 1 } }
