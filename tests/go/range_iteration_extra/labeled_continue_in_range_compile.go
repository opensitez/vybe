// vybe-test: go/range_iteration_extra/labeled_continue_in_range_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { outer: for _, value := range []int{1, 2} { if value == 1 { continue outer } } }
