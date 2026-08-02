// vybe-test: go/go_statement_compile/go_closure_capture_local_string_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { msg := "hi"
go func() { _ = msg }() }
