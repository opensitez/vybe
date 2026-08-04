// vybe-test: go/declarations_patterns/const_block_with_explicit_types_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
const ( a int = 1; b string = "two" )
func main() { _, _ = a, b }
