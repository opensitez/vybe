// vybe-test: go/declarations_patterns/declaration_order_with_types_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
type meter int
var distance meter = 4
func main() { _ = distance }
