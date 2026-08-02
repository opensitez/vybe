// vybe-test: go/go_statement_compile/go_spawn_range_pass_value_arg_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { for _, v := range []int{4, 5} { go func(n int) { _ = n }(v) } }
