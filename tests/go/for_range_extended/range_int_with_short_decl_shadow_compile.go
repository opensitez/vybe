// vybe-test: go/for_range_extended/range_int_with_short_decl_shadow_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 { x := i + 1
_ = x } }
