// vybe-test: go/slice_aliasing_extra/slice_return_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func build() []int { return []int{1, 2} }
func main() { _ = build() }
