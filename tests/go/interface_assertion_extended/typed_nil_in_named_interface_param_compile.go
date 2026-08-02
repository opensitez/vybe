// vybe-test: go/interface_assertion_extended/typed_nil_in_named_interface_param_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
type worker interface { work() }
type task struct{}
func (t *task) work() {}
func accept(w worker) bool { return w == nil }
func main() { var p *task
_ = accept(p) }
