// vybe-test: go/go_statement_compile/go_closure_capture_outer_parameter_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func worker(x int) { go func() { _ = x }() }
func main() { worker(4) }
