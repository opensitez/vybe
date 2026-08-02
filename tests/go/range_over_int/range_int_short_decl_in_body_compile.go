// vybe-test: go/range_over_int/range_int_short_decl_in_body_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 { next := i + 1
_ = next } }
