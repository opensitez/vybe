// vybe-test: go/lang_expressions/take_address_rvalue_compile
// origin: languages/go/tests/go/test_lang_expressions.rs
// vybe-test-mode: compile

package main
func main() { _ = &([]int{1}[0]) }
