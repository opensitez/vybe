// vybe-test: go/go_statement_compile/go_method_on_struct_literal_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type worker struct{}
func (worker) run() {}
func main() { go worker{}.run() }
