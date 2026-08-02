// vybe-test: go/range_over_int/range_int_if_chain_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { for i := range 3 { if i < 1 { _ = i } else if i < 2 { _ = i } else { _ = i } } }
