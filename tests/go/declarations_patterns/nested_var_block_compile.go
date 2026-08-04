// vybe-test: go/declarations_patterns/nested_var_block_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func main() { var a int = 1
{ var b int = a + 1
_ = b }
_ = a }
