// vybe-test: go/slice_aliasing_extra/copy_result_used_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { dst := make([]int, 2)
src := []int{1}
n := copy(dst, src)
_ = n }
