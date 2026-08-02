// vybe-test: go/for_range_extended/range_int_with_defer_in_body_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 { defer func() { _ = i }() } }
