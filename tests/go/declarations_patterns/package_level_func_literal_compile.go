// vybe-test: go/declarations_patterns/package_level_func_literal_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
var op = func(a int, b int) int { return a + b }
func main() { _ = op }
