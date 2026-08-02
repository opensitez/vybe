// vybe-test: go/go_statement_compile/go_method_on_interface_value_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type runner interface { run() }
type worker struct{}
func (worker) run() {}
func main() { var r runner = worker{}
go r.run() }
