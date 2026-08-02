// vybe-test: go/for_range_extended/range_int_switch_per_iteration_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i := range 3 { switch i { case 0: _ = i
case 1: _ = i
default: _ = i } } }
