// vybe-test: go/declarations_patterns/var_block_with_inferred_types_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func main() { var ( a = 1; b = "two"; c = true )
_, _, _ = a, b, c }
