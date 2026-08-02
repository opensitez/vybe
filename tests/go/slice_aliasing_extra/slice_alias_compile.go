// vybe-test: go/slice_aliasing_extra/slice_alias_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { left := []int{1}
right := left
_ = right }
