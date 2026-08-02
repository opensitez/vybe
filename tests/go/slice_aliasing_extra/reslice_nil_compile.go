// vybe-test: go/slice_aliasing_extra/reslice_nil_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { var values []int
_ = values[:] }
