// vybe-test: go/struct_embedding_extra/struct_composite_assignment_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func main() { var left point
left = point{x: 1}
_ = left }
