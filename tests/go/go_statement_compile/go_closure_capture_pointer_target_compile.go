// vybe-test: go/go_statement_compile/go_closure_capture_pointer_target_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { n := 3
p := &n
go func() { _ = *p }() }
