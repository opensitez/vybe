// vybe-test: go/interface_embedding_methods/nil_composite_interface_guarded_call_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type worker interface { work() }
type task interface { worker }
func run(value task) { if value != nil { value.work() } }
func main() { run(nil) }
