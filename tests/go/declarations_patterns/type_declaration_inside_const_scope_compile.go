// vybe-test: go/declarations_patterns/type_declaration_inside_const_scope_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
type score int
const top score = 9
func main() { _ = top }
