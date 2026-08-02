// vybe-test: go/lang_expressions/shift_negative_compile
// origin: languages/go/tests/go/test_lang_expressions.rs
// vybe-test-mode: compile

package main
func main() { var n int
_ = n >> -1 }
