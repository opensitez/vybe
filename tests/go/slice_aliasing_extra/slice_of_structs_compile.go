// vybe-test: go/slice_aliasing_extra/slice_of_structs_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func main() { _ = []point{{x: 1}} }
