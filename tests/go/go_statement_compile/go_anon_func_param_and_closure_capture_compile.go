// vybe-test: go/go_statement_compile/go_anon_func_param_and_closure_capture_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { base := 10
go func(delta int) { _ = base + delta }(3) }
