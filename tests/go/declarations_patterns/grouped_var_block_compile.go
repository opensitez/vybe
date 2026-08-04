// vybe-test: go/declarations_patterns/grouped_var_block_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func main() { var ( a int = 1; b string = "two"; c bool )
_, _, _ = a, b, c }
