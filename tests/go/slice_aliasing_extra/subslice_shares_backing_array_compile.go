// vybe-test: go/slice_aliasing_extra/subslice_shares_backing_array_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1, 2, 3}
part := values[1:]
part[0] = 9
_ = values[1] }
