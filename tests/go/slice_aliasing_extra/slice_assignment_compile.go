// vybe-test: go/slice_aliasing_extra/slice_assignment_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1, 2}
values[0] = 3 }
