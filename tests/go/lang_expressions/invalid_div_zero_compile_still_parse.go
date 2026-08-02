// vybe-test: go/lang_expressions/invalid_div_zero_compile_still_parse
// origin: languages/go/tests/go/test_lang_expressions.rs
// vybe-test-mode: compile

package main
func main() { _ = 1 / 0 }
