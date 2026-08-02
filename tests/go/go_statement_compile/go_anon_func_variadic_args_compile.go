// vybe-test: go/go_statement_compile/go_anon_func_variadic_args_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { go func(nums ...int) { _ = len(nums) }(1, 2, 3) }
