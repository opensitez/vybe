// vybe-test: go/declarations_patterns/grouped_const_block_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
const ( low = 1; high = 9 )
func main() { _ = low
_ = high }
