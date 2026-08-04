// vybe-test: go/declarations_patterns/var_block_mixed_initializers_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func main() { var ( a = 1; b int; c = a + 2 )
_, _, _ = a, b, c }
