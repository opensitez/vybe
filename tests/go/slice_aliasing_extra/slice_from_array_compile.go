// vybe-test: go/slice_aliasing_extra/slice_from_array_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := [3]int{1, 2, 3}
_ = values[1:] }
