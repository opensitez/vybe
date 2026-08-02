// vybe-test: go/struct_embedding_extra/struct_literal_in_return_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func build() point { return point{x: 2} }
func main() { _ = build }
