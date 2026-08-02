// vybe-test: go/functions_patterns_extra/function_returning_pointer_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func build() *point { return &point{x: 1} }
func main() { _ = build() }
