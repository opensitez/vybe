// vybe-test: go/struct_embedding_extra/struct_slice_literal_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func main() { _ = []point{{x: 1}, {x: 2}} }
