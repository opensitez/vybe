// vybe-test: go/slice_aliasing_extra/slicing_in_return_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func build(values []int) []int { return values[1:] }
func main() { _ = build }
