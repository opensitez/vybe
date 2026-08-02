// vybe-test: go/slice_aliasing_extra/make_slice_with_cap_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = make([]int, 2, 4) }
