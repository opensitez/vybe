// vybe-test: go/slice_aliasing_extra/copy_from_nil_slice_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { dst := make([]int, 1)
var src []int
_ = copy(dst, src) }
