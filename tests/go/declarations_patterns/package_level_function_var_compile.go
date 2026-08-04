// vybe-test: go/declarations_patterns/package_level_function_var_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func add(a int, b int) int { return a + b }
var op func(int, int) int = add
func main() { _ = op }
