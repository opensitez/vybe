// vybe-test: go/concurrency_dummy/go_routine_method
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
type Worker struct{}
func (w Worker) Work() {}
func main() { w := Worker{}
go w.Work()
}
