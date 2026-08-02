// vybe-test: go/range_over_int/range_int_with_defer_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { for i := range 3 { defer func() { _ = i }() } }
