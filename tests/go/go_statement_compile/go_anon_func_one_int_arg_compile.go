// vybe-test: go/go_statement_compile/go_anon_func_one_int_arg_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { go func(v int) { _ = v }(7) }
