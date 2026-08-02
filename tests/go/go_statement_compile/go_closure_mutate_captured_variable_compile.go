// vybe-test: go/go_statement_compile/go_closure_mutate_captured_variable_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { n := 0
go func() { n = 5 }() }
