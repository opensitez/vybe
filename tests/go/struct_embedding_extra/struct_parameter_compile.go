// vybe-test: go/struct_embedding_extra/struct_parameter_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func use(value point) int { return value.x }
func main() { _ = use }
