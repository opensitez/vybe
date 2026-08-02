// vybe-test: go/range_over_int/range_int_switch_inside_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { for i := range 4 { switch i { case 0: _ = i
default: _ = i } } }
