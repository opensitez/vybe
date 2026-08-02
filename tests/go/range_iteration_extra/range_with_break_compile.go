// vybe-test: go/range_iteration_extra/range_with_break_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { for _, value := range []int{1} { _ = value
break } }
