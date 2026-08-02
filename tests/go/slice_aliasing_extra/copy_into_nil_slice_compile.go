// vybe-test: go/slice_aliasing_extra/copy_into_nil_slice_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { var dst []int
src := []int{1}
_ = copy(dst, src) }
