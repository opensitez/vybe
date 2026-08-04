// vybe-test: go/declarations_patterns/multiple_init_functions_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
var a int
func init() { a = 1 }
func init() { a = a + 1 }
func main() { _ = a }
