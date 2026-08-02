// vybe-test: go/go_statement_compile/go_anon_func_two_args_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { go func(a int, b string) { _, _ = a, b }(1, "x") }
