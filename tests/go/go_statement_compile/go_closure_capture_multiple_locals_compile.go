// vybe-test: go/go_statement_compile/go_closure_capture_multiple_locals_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { a, b := 1, 2
go func() { _, _ = a, b }() }
