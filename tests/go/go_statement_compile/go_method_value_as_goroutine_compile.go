// vybe-test: go/go_statement_compile/go_method_value_as_goroutine_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type worker struct{}
func (worker) run() {}
func main() { w := worker{}
fn := w.run
go fn() }
