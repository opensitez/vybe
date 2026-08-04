// vybe-test: go/composite_literals_extra/composite_literal_multiline_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int
y int }
func main() { _ = point{
 x: 1,
 y: 2,
 } }
