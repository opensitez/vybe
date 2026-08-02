// vybe-test: go/for_range_extended/range_int_in_function_literal_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { fn := func() { for i := range 2 { _ = i } }
fn() }
