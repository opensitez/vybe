// vybe-test: go/go_statement_compile/go_pointer_receiver_method_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type worker struct { n int }
func (w *worker) bump() { w.n++ }
func main() { w := &worker{}
go w.bump() }
